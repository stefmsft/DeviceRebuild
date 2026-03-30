#Requires -RunAsAdministrator
<#
.SYNOPSIS
    CreatePE.ps1 - Builds and applies customized WinPE environments using the Windows ADK.

.DESCRIPTION
    Two independent phases:

    Phase A — Build PE WIM (no USB required)
      1. copype amd64 C:\WinPE_amd64
      2. Mount boot.wim for customization
      3. Inject PE drivers from PEDrivers\ (optional, multi-select)
      4. Set French keyboard layout (optional)
      5. Add PowerShell support (optional)
      6. Pause for manual customization (optional)
      7. Unmount and commit
      8. Capture staging media as boot.wim in the PE library
         Output: <WinPE_WIM_Location>\<dir>\boot.wim

    Phase B — Apply PE to USB (optional, runs after Phase A or standalone)
      Applies any PE from the library to a USB WinPE partition using
      dism /Apply-Image, then injects the latest startnet.cmd.

    Requires: Windows ADK + WinPE Add-on installed.
              See WINPE_ADK_SETUP.md for installation instructions.

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

$config      = Import-PowerShellDataFile -Path $configPath
$PE_WIM_ROOT = $config["WinPE_WIM_Location"]

if ([string]::IsNullOrWhiteSpace($PE_WIM_ROOT)) {
    Write-Log "WinPE_WIM_Location is not set in config.psd1" -Level Error
    exit 1
}

Write-Log "PE library root: $PE_WIM_ROOT" -Level Info

# ============================================================
# Helper: run a .cmd with ADK environment sourced
# ============================================================
function Invoke-AdkCmd {
    param(
        [string]$CmdFile,
        [string]$Arguments
    )

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
# Phase A: Build PE WIM (optional — no USB required)
# ============================================================
Write-Section "Phase A: Build PE WIM"

Write-Host ""
Write-Log "Builds a customized WinPE using the ADK and saves it to the PE library." -Level Info
Write-Log "No USB key required for this phase." -Level Info
Write-Host ""

$builtWimPath = $null   # set if Phase A completes successfully
$builtDirName = $null

if (Read-YesNo "Build or update a PE WIM?") {

    # ----------------------------------------------------------
    # Select output directory in PE library
    # ----------------------------------------------------------
    Write-Section "Select Output Directory"

    Write-Log "Select where to save the new boot.wim in the PE library." -Level Info

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
        Write-Host "Enter new directory name: " -NoNewline -ForegroundColor Yellow
        $dirName = (Read-Host).Trim()
        if ([string]::IsNullOrWhiteSpace($dirName)) {
            Write-Log "No directory name provided. Aborting." -Level Error
            exit 1
        }
        Write-Log "New directory: $dirName" -Level Info
    } else {
        $dirName     = $selectedOption
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

    # ----------------------------------------------------------
    # Build steps
    # ----------------------------------------------------------
    $mountActive = $false

    try {

        # ======================================================
        # Step 1: copype
        # ======================================================
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

        # ======================================================
        # Step 2: Mount boot.wim
        # ======================================================
        Write-Section "Step 2: Mount boot.wim"

        Write-Log "Mounting: $bootWim" -Level Info
        Write-Host ""

        & dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            Write-Log "Failed to mount boot.wim (exit code: $dismExitCode)" -Level Error
            exit 1
        }

        Write-Log "boot.wim mounted successfully." -Level Success
        $mountActive = $true

        # ======================================================
        # Step 3: PE Driver Injection (Optional)
        # ======================================================
        Write-Section "Step 3: PE Driver Injection (Optional)"

        $peDriversRoot = Join-Path $PE_WIM_ROOT "PEDrivers"

        if (-not (Test-Path $peDriversRoot -PathType Container)) {
            Write-Host ""
            Write-Log "No PEDrivers folder found at: $peDriversRoot" -Level Info
            Write-Log "Create PEDrivers\ with driver subfolders to enable this step." -Level Info
        } else {
            $driverDirs = @(Get-ChildItem -Path $peDriversRoot -Directory |
                            Select-Object -ExpandProperty Name)

            if ($driverDirs.Count -eq 0) {
                Write-Host ""
                Write-Log "PEDrivers folder is empty — no driver directories to inject." -Level Warning
            } else {
                Write-Host ""
                Write-Log "Select driver packages to inject into WinPE:" -Level Info

                $selectedNames = Select-MultiFromList -Items $driverDirs `
                                                      -Prompt "Toggle (number), A=all, or Enter to confirm"

                if ($selectedNames.Count -eq 0) {
                    Write-Log "No drivers selected — injection skipped." -Level Info
                } else {
                    foreach ($drvName in $selectedNames) {
                        $drvPath = Join-Path $peDriversRoot $drvName
                        Write-Host ""
                        Write-Log "Injecting: $drvName" -Level Info
                        Write-Log "(This may take several minutes — please wait)" -Level Warning
                        Write-Host ""
                        & dism.exe /Image:"$mountDir" /Add-Driver "/Driver:$drvPath" /Recurse
                        $dismExitCode = $LASTEXITCODE

                        if ($dismExitCode -ne 0) {
                            Write-Log "Warning: Some drivers may have failed for '$drvName' (exit code: $dismExitCode)" -Level Warning
                        } else {
                            Write-Log "'$drvName' injected successfully." -Level Success
                        }
                    }
                }
            }
        }

        # ======================================================
        # Step 4: French keyboard layout (optional)
        # ======================================================
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

        # ======================================================
        # Step 5: PowerShell support (optional)
        # ======================================================
        Write-Section "Step 5: PowerShell Support (Optional)"

        Write-Host ""
        Write-Log "Adds PowerShell 5.1 to WinPE. Installs 6 components in order:" -Level Info
        Write-Log "  WinPE-WMI, WinPE-NetFX, WinPE-Scripting, WinPE-PowerShell," -Level Info
        Write-Log "  WinPE-StorageWMI, WinPE-DismCmdlets" -Level Info
        Write-Log "Size impact: ~120 MB (mostly .NET runtime)." -Level Warning
        Write-Host ""

        if (Read-YesNo "Add PowerShell support to WinPE?") {

            $ocRoot = Join-Path $adkRoot "Windows Preinstallation Environment\amd64\WinPE_OCs"

            $psPackages = @(
                @( "WinPE-WMI.cab",          "en-us\WinPE-WMI_en-us.cab"          ),
                @( "WinPE-NetFX.cab",         "en-us\WinPE-NetFX_en-us.cab"         ),
                @( "WinPE-Scripting.cab",     "en-us\WinPE-Scripting_en-us.cab"     ),
                @( "WinPE-PowerShell.cab",    "en-us\WinPE-PowerShell_en-us.cab"    ),
                @( "WinPE-StorageWMI.cab",    "en-us\WinPE-StorageWMI_en-us.cab"    ),
                @( "WinPE-DismCmdlets.cab",   "en-us\WinPE-DismCmdlets_en-us.cab"   )
            )

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

        # ======================================================
        # Step 6: Manual customization pause
        # ======================================================
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

        # ======================================================
        # Step 7: Unmount and commit
        # ======================================================
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

        # ======================================================
        # Step 8: Capture staging media as PE library WIM
        # ======================================================
        Write-Section "Step 8: Capture boot.wim to PE Library"

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
        } else {
            $finalSizeMB = [math]::Round((Get-Item $destWim).Length / 1MB, 0)
            Write-Log "WIM captured: $destWim  (${finalSizeMB} MB)" -Level Success
            $builtWimPath = $destWim
            $builtDirName = $dirName
        }

    }
    finally {
        if ($mountActive) {
            Write-Log "Discarding stale WIM mount..." -Level Warning
            & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
        }

        if (Test-Path $stagingDir) {
            Write-Log "Removing staging directory..." -Level Info
            $prog = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
            Remove-Item $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
            $ProgressPreference = $prog
            Write-Log "Staging directory removed." -Level Info
        }
    }

} else {
    Write-Log "Build step skipped." -Level Info
}

# ============================================================
# Phase B: Apply PE to USB Partition (optional)
# ============================================================
Write-Section "Phase B: Apply PE to USB Partition"

Write-Host ""
Write-Log "Applies a WinPE image from the library to a USB WinPE partition." -Level Info
Write-Log "Works with any PE in the library — including one just built above." -Level Info
Write-Host ""

if (Read-YesNo "Apply a PE to a USB WinPE partition?") {

    # Collect available PEs from library
    $peEntries = @()
    if (Test-Path $PE_WIM_ROOT -PathType Container) {
        $peEntries = @(Get-ChildItem -Path $PE_WIM_ROOT -Directory |
                       Where-Object { Test-Path (Join-Path $_.FullName "boot.wim") } |
                       Select-Object -ExpandProperty Name)
    }

    if ($peEntries.Count -eq 0) {
        Write-Log "No boot.wim found in PE library: $PE_WIM_ROOT" -Level Warning
        Write-Log "Build a PE first using Phase A." -Level Warning
    } else {

        if ($builtDirName) {
            Write-Log "Just built: $builtDirName  (available in list below)" -Level Info
        }

        Write-Host ""
        $selectedPEName = Select-FromList -Items $peEntries -Prompt "Select PE to apply"
        $selectedWim    = Join-Path $PE_WIM_ROOT "$selectedPEName\boot.wim"

        Write-Host ""
        Write-Log "Select the WinPE partition on the target USB key (FAT32, typically P:)." -Level Info
        Write-Log "WARNING: All existing content on this partition will be replaced." -Level Warning

        $selectedDrive = Select-DriveLetter -Purpose "WinPE partition"

        if (-not $selectedDrive) {
            Write-Log "No drive selected. Skipping." -Level Warning
        } else {
            $usbLetter = $selectedDrive.DriveLetter
            Write-Host ""
            if (Read-YesNo "Apply '$selectedPEName' to ${usbLetter}? All content will be replaced.") {

                Write-Log "Applying WinPE image to $usbLetter..." -Level Info
                Write-Host ""
                & dism.exe /Apply-Image "/ImageFile:$selectedWim" /Index:1 "/ApplyDir:$usbLetter"
                $dismExitCode = $LASTEXITCODE

                if ($dismExitCode -ne 0) {
                    Write-Log "DISM apply failed (exit code: $dismExitCode)" -Level Error
                } else {
                    Write-Log "WinPE image applied to $usbLetter." -Level Success

                    # Inject latest startnet.cmd from Scripts\
                    $startnetOk = Update-StartnetCmd -WinPEDriveLetter $usbLetter
                    if (-not $startnetOk) {
                        Write-Log "Warning: startnet.cmd could not be updated in boot.wim." -Level Warning
                    }

                    Write-Log "WinPE partition $usbLetter is ready." -Level Success
                }
            } else {
                Write-Log "Apply cancelled by user." -Level Warning
            }
        }
    }

} else {
    Write-Log "Apply step skipped." -Level Info
}

# ============================================================
# Summary
# ============================================================
Write-Banner "Complete"

Write-Log "CreatePE finished." -Level Success
Write-Host ""
if ($builtWimPath) {
    Write-Host "  PE library entry : $builtWimPath" -ForegroundColor White
}
Write-Host ""
Write-Host "  Use ProduceKey.ps1 to apply a PE to a USB key and copy model WIMs." -ForegroundColor Cyan
Write-Host ""

if ($Script:LogEnabled) {
    Write-Log "Log saved to: $Script:LogFile" -Level Info
}
