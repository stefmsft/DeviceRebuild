#Requires -RunAsAdministrator
<#
.SYNOPSIS
    CreatePE.ps1 - Builds a customized WinPE environment using the Windows ADK.

.DESCRIPTION
    Uses the ADK's copype command to create a proper WinPE working directory,
    optionally injects Intel RST/storage drivers, captures the result as a
    boot.wim usable by ProduceKey.ps1, then writes a bootable USB key via
    MakeWinPEMedia.

    Workflow:
      1. copype amd64 C:\WinPE_amd64
      2. Mount boot.wim for customization
      3. Inject iRST/storage drivers (optional)
      4. Pause for manual customization (optional)
      5. Unmount and commit
      6. Capture staging media as ProduceKey-compatible boot.wim
      7. MakeWinPEMedia /UFD C:\WinPE_amd64 <USB-drive>
      8. Cleanup

    Requires: Windows ADK + WinPE Add-on installed.
              See WINPE_ADK_SETUP.md for installation instructions.

    Output:  <WinPE_WIM_Location>\<Selected-Dir>\boot.wim

.PARAMETER Log
    Enable logging to file.
#>
param(
    [switch]$Log
)

# ============================================================
# Module + logging init
# ============================================================
Import-Module (Join-Path $PSScriptRoot "DeviceRebuild.psm1") -Force

$Script:LogEnabled = $Log.IsPresent
$Script:LogFile    = if ($Script:LogEnabled) { "CreatePE_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" } else { $null }
Initialize-Logging -Enabled $Script:LogEnabled -LogFile $Script:LogFile

$stagingDir = "C:\WinPE_amd64"
$mountDir   = Join-Path $stagingDir "mount"
$mediaDir   = Join-Path $stagingDir "media"
$bootWim    = Join-Path $mediaDir "sources\boot.wim"

# ============================================================
# ADK Check
# ============================================================
Write-Banner "CreatePE - WinPE Builder"

$adkRoot        = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"
$dandISetEnv    = Join-Path $adkRoot "Deployment Tools\DandISetEnv.bat"
$copypeCmd      = Join-Path $adkRoot "Windows Preinstallation Environment\copype.cmd"
$makeWinPEMedia = Join-Path $adkRoot "Windows Preinstallation Environment\MakeWinPEMedia.cmd"

$adkOk = (Test-Path $copypeCmd) -and (Test-Path $makeWinPEMedia)

if (-not $adkOk) {
    Write-Log "Windows ADK + WinPE Add-on not found." -Level Error
    Write-Log "" -Level Error
    Write-Log "This script requires:" -Level Error
    Write-Log "  1. Windows Assessment and Deployment Kit (ADK)" -Level Error
    Write-Log "  2. Windows PE add-on for the ADK" -Level Error
    Write-Log "" -Level Error
    Write-Log "Installation instructions: $PSScriptRoot\WINPE_ADK_SETUP.md" -Level Error
    Write-Log "" -Level Error
    Write-Log "Expected files after installation:" -Level Error
    Write-Log "  $copypeCmd" -Level Error
    Write-Log "  $makeWinPEMedia" -Level Error
    exit 1
}

Write-Log "ADK found." -Level Success

# ============================================================
# Load configuration
# ============================================================
Write-Section "Load Configuration"

$configPath = Join-Path $PSScriptRoot "config.psd1"
if (-not (Test-Path $configPath)) {
    Write-Log "config.psd1 not found: $configPath" -Level Error
    exit 1
}

$config          = Import-PowerShellDataFile -Path $configPath
$PE_WIM_ROOT     = $config["WinPE_WIM_Location"]
$DEV_ROOT_WIM    = $config["Device_Root_WIM_Location"]
$SHARED_DRV_ROOT = $config["Shared_Drivers_Location"]

if ([string]::IsNullOrWhiteSpace($PE_WIM_ROOT)) {
    Write-Log "WinPE_WIM_Location is not set in config.psd1" -Level Error
    exit 1
}

Write-Log "WinPE output root : $PE_WIM_ROOT" -Level Info
if ($DEV_ROOT_WIM)    { Write-Log "Device WIM root   : $DEV_ROOT_WIM" -Level Info }
if ($SHARED_DRV_ROOT) { Write-Log "Shared drivers    : $SHARED_DRV_ROOT" -Level Info }

# ============================================================
# Select output directory (in WinPE_WIM_Location)
# ============================================================
Write-Section "Select Output Directory"

Write-Log "Select where to save the boot.wim for ProduceKey." -Level Info

$existingDirs = @()
if (Test-Path $PE_WIM_ROOT -PathType Container) {
    $existingDirs = @(Get-ChildItem -Path $PE_WIM_ROOT -Directory |
                      Select-Object -ExpandProperty Name)
}

$newDirOption = "[+ New directory]"
$dirOptions   = @($existingDirs) + @($newDirOption)

$selectedOption = Select-FromList -Items $dirOptions -Prompt "Select output directory"

if ($selectedOption -eq $newDirOption) {
    Write-Host ""
    Write-Host "Enter new directory name (spaces allowed): " -NoNewline -ForegroundColor Yellow
    $dirName = Read-Host
    $dirName  = $dirName.Trim()
    if ([string]::IsNullOrWhiteSpace($dirName)) {
        Write-Log "No directory name provided. Aborting." -Level Error
        exit 1
    }
    Write-Log "New directory: $dirName" -Level Info
} else {
    $dirName = $selectedOption
    $existingWim = Join-Path (Join-Path $PE_WIM_ROOT $dirName) "boot.wim"
    Write-Host ""
    Write-Log "WARNING: '$dirName' already exists." -Level Warning
    if (Test-Path $existingWim) {
        Write-Log "         Existing boot.wim will be deleted and replaced." -Level Warning
    }
    Write-Host ""
    if (-not (Read-YesNo "Continue and overwrite '$dirName'?")) {
        Write-Log "Cancelled by user." -Level Warning
        exit 0
    }
}

$destDir = Join-Path $PE_WIM_ROOT $dirName
$destWim = Join-Path $destDir "boot.wim"

Write-Log "Output WIM: $destWim" -Level Info

# ============================================================
# Select USB target partition
# ============================================================
Write-Section "Select USB Target Partition"

Write-Log "Select the partition that MakeWinPEMedia will write the bootable WinPE to." -Level Info
Write-Log "WARNING: MakeWinPEMedia will overwrite all content on the selected partition." -Level Warning

$selectedDrive = Select-DriveLetter -Purpose "MakeWinPEMedia target"
if (-not $selectedDrive) {
    Write-Log "No drive selected. Aborting." -Level Error
    exit 1
}

$usbLetter = $selectedDrive.DriveLetter
Write-Host ""
if (-not (Read-YesNo "All content on $usbLetter will be replaced. Continue?")) {
    Write-Log "Cancelled by user." -Level Warning
    exit 0
}

# ============================================================
# Helper: run a .cmd with ADK environment sourced
# ============================================================
function Invoke-AdkCmd {
    param(
        [string]$CmdFile,
        [string]$Arguments
    )

    # Write a temp batch file to sidestep PowerShell/cmd quoting issues with
    # paths containing spaces (e.g. "C:\Program Files (x86)\...").
    $tempBat = Join-Path $env:TEMP "CreatePE_$(Get-Random).bat"
    try {
        $lines = @('@echo off')
        if (Test-Path $dandISetEnv) {
            $lines += "call `"$dandISetEnv`""
        }
        $lines += "call `"$CmdFile`" $Arguments"
        $lines | Set-Content $tempBat -Encoding ASCII

        $p = Start-Process -FilePath "cmd.exe" `
            -ArgumentList "/c `"$tempBat`"" `
            -NoNewWindow -Wait -PassThru
        return $p.ExitCode
    } finally {
        Remove-Item $tempBat -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
# Step 1: copype
# ============================================================
Write-Section "Step 1: copype - Create WinPE Working Directory"

if (Test-Path $stagingDir) {
    Write-Log "Removing existing staging directory: $stagingDir" -Level Info
    $prog = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
    Remove-Item $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
    $ProgressPreference = $prog
}

Write-Log "Running: copype amd64 $stagingDir" -Level Info
Write-Host ""

$exitCode = Invoke-AdkCmd -CmdFile $copypeCmd -Arguments "amd64 `"$stagingDir`""

if ($exitCode -ne 0) {
    Write-Log "copype failed (exit code: $exitCode)" -Level Error
    exit 1
}

if (-not (Test-Path $bootWim)) {
    Write-Log "copype completed but boot.wim was not found at: $bootWim" -Level Error
    exit 1
}

Write-Log "WinPE working directory ready." -Level Success

# ============================================================
# Step 2: Mount boot.wim
# ============================================================
Write-Section "Step 2: Mount boot.wim"

Write-Log "Mounting: $bootWim" -Level Info
Write-Log "Mount directory: $mountDir" -Level Info
Write-Host ""

& dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
$dismExitCode = $LASTEXITCODE

if ($dismExitCode -ne 0) {
    Write-Log "Failed to mount boot.wim (exit code: $dismExitCode)" -Level Error
    # Cleanup staging before exiting
    $prog = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
    Remove-Item $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
    $ProgressPreference = $prog
    exit 1
}

Write-Log "boot.wim mounted successfully." -Level Success

# From here we must ensure unmount happens — use try/finally
$mountActive = $true

try {

    # ============================================================
    # Step 3: iRST Driver Injection (Optional)
    # ============================================================
    Write-Section "Step 3: iRST Driver Injection (Optional)"

    Write-Host ""
    Write-Log "Injecting Intel RST/storage drivers ensures WinPE can detect NVMe/SATA" -Level Info
    Write-Log "controllers on devices that do not have inbox driver support." -Level Info
    Write-Host ""

    if (Read-YesNo "Inject iRST drivers into WinPE?") {

        $rstEntries = @()

        # --- Scan Device WIM repository for models with Drivers\Rapid Storage ---
        if (-not [string]::IsNullOrWhiteSpace($DEV_ROOT_WIM) -and
            (Test-Path $DEV_ROOT_WIM -PathType Container)) {

            Write-Log "Scanning WIM repository for RST driver sources..." -Level Info

            foreach ($modelDir in (Get-ChildItem $DEV_ROOT_WIM -Directory)) {
                $srcDriverDir = Join-Path $modelDir.FullName "Drivers"
                if (-not (Test-Path $srcDriverDir -PathType Container)) { continue }

                # Follow LNK redirect if Shared_Drivers_Location is configured
                if (-not [string]::IsNullOrWhiteSpace($SHARED_DRV_ROOT)) {
                    $driverResolved = Resolve-DriverLnk -DriversDir $srcDriverDir -SharedDriversLocation $SHARED_DRV_ROOT
                } else {
                    $driverResolved = [PSCustomObject]@{ Path = $srcDriverDir; LinkedModel = $null }
                }

                if (-not $driverResolved) { continue }

                $rst = Find-RSTDriverPath -DriverDir $driverResolved.Path
                if (-not $rst) { continue }

                $lnkSuffix  = if ($driverResolved.LinkedModel) { "  [->$($driverResolved.LinkedModel)]" } else { '' }
                $rstEntries += [PSCustomObject]@{
                    Label = "$($modelDir.Name)$lnkSuffix"
                    Path  = $rst
                }
            }
        }

        # --- Also scan Shared_Drivers_Location top-level for direct RST entries ---
        if (-not [string]::IsNullOrWhiteSpace($SHARED_DRV_ROOT) -and
            (Test-Path $SHARED_DRV_ROOT -PathType Container)) {

            foreach ($sharedDir in (Get-ChildItem $SHARED_DRV_ROOT -Directory)) {
                $rst = Find-RSTDriverPath -DriverDir $sharedDir.FullName
                if (-not $rst) { continue }

                # Skip if the same path was already discovered via the repo scan
                $alreadyListed = $rstEntries | Where-Object { $_.Path -ieq $rst }
                if ($alreadyListed) { continue }

                $rstEntries += [PSCustomObject]@{
                    Label = "$($sharedDir.Name)  [shared drivers]"
                    Path  = $rst
                }
            }
        }

        if ($rstEntries.Count -eq 0) {
            Write-Log "No RST/storage driver sources found. Skipping injection." -Level Warning
            Write-Log "Expected structure: <WIM repo>\<Model>\Drivers\Rapid Storage\" -Level Warning
        } else {
            Write-Host ""
            $msg = if ($rstEntries.Count -eq 1) { 'Driver source found' } else { 'Multiple driver sources found' }
            Write-Log "$msg — select one to inject:" -Level Info

            $selectedLabel = Select-FromList -Items ($rstEntries | Select-Object -ExpandProperty Label) `
                                            -Prompt "Select driver source"
            $selectedEntry = $rstEntries | Where-Object { $_.Label -eq $selectedLabel }

            if ($selectedEntry) {
                Write-Log "Injecting drivers from: $($selectedEntry.Path)" -Level Info
                Write-Log "(Driver injection may take several minutes — please wait)" -Level Warning
                Write-Host ""
                & dism.exe /Image:"$mountDir" /Add-Driver "/Driver:$($selectedEntry.Path)" /Recurse
                $dismExitCode = $LASTEXITCODE

                if ($dismExitCode -ne 0) {
                    Write-Log "Warning: Some drivers may have failed (exit code: $dismExitCode)" -Level Warning
                } else {
                    Write-Log "Drivers injected successfully." -Level Success
                }
            }
        }

    } else {
        Write-Log "Driver injection skipped." -Level Info
    }

    # ============================================================
    # Step 4: French keyboard layout (optional)
    # ============================================================
    Write-Section "Step 4: French Keyboard Layout (fr-FR)"

    Write-Host ""
    Write-Log "Sets the default keyboard to AZERTY (fr-FR). No CAB files required." -Level Info
    Write-Host ""

    if (Read-YesNo "Set French keyboard layout (fr-FR / AZERTY)?") {
        & dism.exe /Image:"$mountDir" /Set-InputLocale:fr-FR
        $dismExitCode = $LASTEXITCODE
        if ($dismExitCode -ne 0) {
            Write-Log "Warning: Set-InputLocale failed (exit code: $dismExitCode)" -Level Warning
        } else {
            Write-Log "Keyboard layout set to fr-FR (AZERTY)." -Level Success
        }
    } else {
        Write-Log "French keyboard layout skipped." -Level Info
    }

    # ============================================================
    # Step 5: PowerShell support (optional)
    # ============================================================
    Write-Section "Step 5: PowerShell Support (Optional)"

    Write-Host ""
    Write-Log "Adds PowerShell 5.1 to WinPE. Installs 6 components in order:" -Level Info
    Write-Log "  WinPE-WMI, WinPE-NetFX, WinPE-Scripting, WinPE-PowerShell," -Level Info
    Write-Log "  WinPE-StorageWMI, WinPE-DismCmdlets" -Level Info
    Write-Log "Size impact: ~120 MB (mostly .NET runtime)." -Level Warning
    Write-Host ""

    if (Read-YesNo "Add PowerShell support to WinPE?") {

        $ocRoot = Join-Path $adkRoot "Windows Preinstallation Environment\amd64\WinPE_OCs"

        # Ordered list: [language-neutral, language-specific] pairs
        $psPackages = @(
            @( "WinPE-WMI.cab",          "en-us\WinPE-WMI_en-us.cab"          ),
            @( "WinPE-NetFX.cab",         "en-us\WinPE-NetFX_en-us.cab"         ),
            @( "WinPE-Scripting.cab",     "en-us\WinPE-Scripting_en-us.cab"     ),
            @( "WinPE-PowerShell.cab",    "en-us\WinPE-PowerShell_en-us.cab"    ),
            @( "WinPE-StorageWMI.cab",    "en-us\WinPE-StorageWMI_en-us.cab"    ),
            @( "WinPE-DismCmdlets.cab",   "en-us\WinPE-DismCmdlets_en-us.cab"   )
        )

        # Pre-flight: verify every CAB exists before touching the image
        $missingCabs = @()
        foreach ($pair in $psPackages) {
            foreach ($rel in $pair) {
                $full = Join-Path $ocRoot $rel
                if (-not (Test-Path $full)) { $missingCabs += $full }
            }
        }

        if ($missingCabs.Count -gt 0) {
            Write-Log "Cannot add PowerShell — the following CAB files were not found:" -Level Error
            foreach ($m in $missingCabs) { Write-Log "  $m" -Level Error }
            Write-Log "Ensure the WinPE add-on for the ADK is fully installed." -Level Warning
        } else {
            $psOk = $true
            foreach ($pair in $psPackages) {
                foreach ($rel in $pair) {
                    $cabPath = Join-Path $ocRoot $rel
                    Write-Log "Installing: $rel" -Level Info
                    Write-Host ""
                    & dism.exe /Add-Package "/Image:$mountDir" "/PackagePath:$cabPath"
                    if ($LASTEXITCODE -ne 0) {
                        Write-Log "Warning: package install returned exit code $LASTEXITCODE for $rel" -Level Warning
                        $psOk = $false
                    }
                }
            }

            if ($psOk) {
                Write-Log "PowerShell support added successfully." -Level Success
            } else {
                Write-Log "PowerShell support installed with warnings — check output above." -Level Warning
            }
        }

    } else {
        Write-Log "PowerShell support skipped." -Level Info
    }

    # ============================================================
    # Step 6: Manual customization pause (before unmount)
    # ============================================================
    Write-Section "Step 6: Additional Customization"

    Write-Host ""
    Write-Log "The WinPE image is currently mounted at:" -Level Info
    Write-Host "  $mountDir" -ForegroundColor Cyan
    Write-Host ""
    Write-Log "You can make additional changes before the image is committed:" -Level Info
    Write-Log "  - Edit Windows\System32\startnet.cmd" -Level Info
    Write-Log "  - Add scripts or tools" -Level Info
    Write-Log "  - Copy any other files you need inside WinPE" -Level Info
    Write-Host ""

    if (Read-YesNo "Pause here for manual customization?") {
        Write-Host ""
        Write-Host "  Mount path : " -NoNewline -ForegroundColor Yellow
        Write-Host $mountDir -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Make your changes in the folder above," -ForegroundColor Yellow
        Write-Host "  then press Enter to unmount and commit." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to continue"
        Write-Log "Resuming after manual customization." -Level Info
    }

    # ============================================================
    # Step 5: Unmount and commit
    # ============================================================
    Write-Section "Step 7: Unmount and Commit"

    Write-Log "Committing changes and unmounting..." -Level Info
    Write-Host ""

    & dism.exe /Unmount-Image "/MountDir:$mountDir" /Commit
    $dismExitCode = $LASTEXITCODE

    if ($dismExitCode -ne 0) {
        Write-Log "Unmount/commit failed (exit code: $dismExitCode)" -Level Error
        Write-Log "Attempting to discard changes..." -Level Warning
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
        $mountActive = $false
        exit 1
    }

    $mountActive = $false
    Write-Log "Changes committed." -Level Success

    # ============================================================
    # Step 6: Capture staging media as ProduceKey-compatible WIM
    # ============================================================
    Write-Section "Step 8: Capture boot.wim for ProduceKey"

    if (-not (Test-Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        Write-Log "Created directory: $destDir" -Level Info
    }

    if (Test-Path $destWim) {
        Remove-Item $destWim -Force
        Write-Log "Removed existing boot.wim." -Level Info
    }

    Write-Log "Capturing $mediaDir" -Level Info
    Write-Log "      to  $destWim" -Level Info
    Write-Log "(This may take several minutes)" -Level Warning
    Write-Host ""

    & dism.exe /Capture-Image `
        "/ImageFile:$destWim" `
        "/CaptureDir:$mediaDir" `
        "/Name:WinPE-$dirName" `
        /Compress:max

    $captureExitCode = $LASTEXITCODE

    if ($captureExitCode -ne 0) {
        Write-Log "Capture failed (exit code: $captureExitCode)" -Level Error
        Write-Log "The USB will still be written. Manually re-run the capture if needed." -Level Warning
    } else {
        $finalSizeMB = [math]::Round((Get-Item $destWim).Length / 1MB, 0)
        Write-Log "WIM captured: $destWim  (${finalSizeMB} MB)" -Level Success
    }

    # ============================================================
    # Step 7: MakeWinPEMedia — write bootable USB
    # ============================================================
    Write-Section "Step 9: MakeWinPEMedia - Write Bootable USB"

    Write-Log "Writing bootable WinPE to $usbLetter ..." -Level Info
    Write-Host ""

    $exitCode = Invoke-AdkCmd -CmdFile $makeWinPEMedia -Arguments "/UFD `"$stagingDir`" $usbLetter"

    if ($exitCode -ne 0) {
        Write-Log "MakeWinPEMedia reported an error (exit code: $exitCode)" -Level Error
        Write-Log "The boot.wim for ProduceKey was already saved — USB step only failed." -Level Warning
    } else {
        Write-Log "Bootable WinPE USB ready at $usbLetter" -Level Success
    }

}
finally {
    # Safety net: discard stale mount if an error interrupted the unmount step
    if ($mountActive) {
        Write-Log "Discarding stale WIM mount..." -Level Warning
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
    }

    # Always clean up staging directory
    if (Test-Path $stagingDir) {
        Write-Log "Removing staging directory..." -Level Info
        $prog = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        Remove-Item $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
        $ProgressPreference = $prog
        Write-Log "Staging directory removed." -Level Info
    }
}

# ============================================================
# Summary
# ============================================================
Write-Banner "Complete"

Write-Log "CreatePE finished." -Level Success
Write-Host ""
Write-Host "  Results:" -ForegroundColor Cyan
Write-Host "    Bootable USB    : $usbLetter" -ForegroundColor White
Write-Host "    ProduceKey WIM  : $destWim" -ForegroundColor White
Write-Host ""
Write-Host "  '$dirName\boot.wim' is now available in ProduceKey.ps1." -ForegroundColor Green
Write-Host ""

if ($Script:LogEnabled) {
    Write-Log "Log saved to: $Script:LogFile" -Level Info
}
