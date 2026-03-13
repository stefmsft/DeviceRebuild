#Requires -RunAsAdministrator
<#
.SYNOPSIS
    ProduceKey.ps1 - USB Key Creation Tool for Windows Device Imaging

.DESCRIPTION
    Creates bootable USB keys for capturing and deploying Windows images.
    Supports dual-partition USB (WinPE boot + NTFS storage).

.PARAMETER Log
    Enable logging to file. By default, no log file is created.

.NOTES
    Requires: Administrator privileges
#>
param(
    [switch]$Log
)

# ============================================================
# Module Import
# ============================================================
Import-Module (Join-Path $PSScriptRoot "DeviceRebuild.psm1") -Force

# ============================================================
# Configuration
# ============================================================
$Script:SizeLimit  = 128GB
$Script:LogEnabled = $Log.IsPresent
$Script:LogFile    = if ($Script:LogEnabled) { "ProduceKey_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" } else { $null }

Initialize-Logging -Enabled $Script:LogEnabled -LogFile $Script:LogFile

# ============================================================
# Core Operations
# ============================================================
# Note: Get-USBDrives, Show-USBDrives, Select-DriveLetter are in DeviceRebuild.psm1

function Select-USBDisk {
    <#
    .SYNOPSIS
        Prompts user to select a USB disk for formatting
    #>

    $usbDisks = Get-Disk | Where-Object { $_.BusType -eq 'USB' -and $_.Size -gt 0 }

    if ($usbDisks.Count -eq 0) {
        Write-Log "No USB storage devices found." -Level Error
        return $null
    }

    Write-Host ""
    Write-Host "  Available USB Disks:" -ForegroundColor Cyan
    Write-Host "  " + ("-" * 56) -ForegroundColor DarkGray

    $diskList = @()
    foreach ($disk in $usbDisks) {
        $sizeGB = [math]::Round($disk.Size / 1GB, 2)
        $diskList += [PSCustomObject]@{
            Number = $disk.Number
            Name   = $disk.FriendlyName
            SizeGB = $sizeGB
            Display = "Disk $($disk.Number): $($disk.FriendlyName) - ${sizeGB}GB"
        }
        Write-Host "  [$($disk.Number)] $($disk.FriendlyName) - ${sizeGB}GB" -ForegroundColor White
    }

    Write-Host "  " + ("-" * 56) -ForegroundColor DarkGray
    Write-Host ""

    do {
        Write-Host "Enter disk number to format (or 'Q' to cancel): " -NoNewline -ForegroundColor Yellow
        $input = Read-Host

        if ($input -match '^[Qq]') {
            return $null
        }

        $diskNumber = $input -as [int]
        $selectedDisk = $diskList | Where-Object { $_.Number -eq $diskNumber }

        if (-not $selectedDisk) {
            Write-Log "Invalid disk number. Please try again." -Level Warning
        }

    } while (-not $selectedDisk)

    # Safety check - size limit
    if ($selectedDisk.SizeGB -gt ($Script:SizeLimit / 1GB)) {
        Write-Log "Disk exceeds size limit of $($Script:SizeLimit / 1GB)GB. Aborting for safety." -Level Error
        return $null
    }

    return $selectedDisk
}

function Format-USBDrive {
    <#
    .SYNOPSIS
        Formats a USB drive with dual partitions (WinPE FAT32 + Images NTFS)
    #>

    Write-Section "Format USB Drive"

    $disk = Select-USBDisk

    if (-not $disk) {
        return $null
    }

    Write-Host ""
    Write-Host "  WARNING: All data on Disk $($disk.Number) will be erased!" -ForegroundColor Red
    Write-Host "  Device: $($disk.Name)" -ForegroundColor Red
    Write-Host "  Size: $($disk.SizeGB)GB" -ForegroundColor Red
    Write-Host ""

    if (-not (Read-YesNo "Are you sure you want to continue?")) {
        Write-Log "Format cancelled by user." -Level Warning
        return $null
    }

    Write-Log "Formatting Disk $($disk.Number)..." -Level Info

    # Generate diskpart script
    $diskpartScript = @"
select disk $($disk.Number)
clean
convert mbr
create partition primary size=3072
select partition 1
active
format fs=FAT32 quick label="WinPE"
assign letter=P
create partition primary
format fs=NTFS quick label="Images"
assign letter=I
list partition
exit
"@

    $scriptPath = Join-Path $env:TEMP "PrepareUSB_$(Get-Random).txt"
    $diskpartScript | Out-File -FilePath $scriptPath -Encoding ASCII

    Write-Log "Running diskpart..." -Level Info
    $process = Start-Process -FilePath "diskpart.exe" -ArgumentList "/s `"$scriptPath`"" -NoNewWindow -Wait -PassThru

    # Log diskpart output
    if ($Script:LogEnabled) {
        Add-Content -Path $Script:LogFile -Value "Diskpart script executed: $scriptPath"
    }

    Remove-Item -Path $scriptPath -Force -ErrorAction SilentlyContinue

    if ($process.ExitCode -eq 0) {
        Write-Log "Disk formatted successfully." -Level Success
        Write-Log "  WinPE partition: P: (3GB FAT32)" -Level Info
        Write-Log "  Images partition: I: (NTFS)" -Level Info

        return [PSCustomObject]@{
            WinPELetter  = "P:"
            ImagesLetter = "I:"
            DiskNumber   = $disk.Number
        }
    } else {
        Write-Log "Diskpart failed with exit code: $($process.ExitCode)" -Level Error
        return $null
    }
}

function Install-WinPE {
    <#
    .SYNOPSIS
        Applies a WinPE WIM file to a partition using DISM /Apply-Image.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WimFile,

        [Parameter(Mandatory)]
        [string]$DriveLetter
    )

    Write-Section "Install WinPE"

    if (-not (Test-Path $WimFile)) {
        Write-Log "WinPE WIM file not found: $WimFile" -Level Error
        return $false
    }

    Write-Log "Applying WinPE to $DriveLetter..." -Level Info
    Write-Log "  Source: $WimFile" -Level Info

    try {
        # Use call operator to avoid Start-Process -Wait hanging on DismHost.exe
        # child processes that outlive dism.exe itself.
        Write-Host ""
        & dism.exe /Apply-Image "/ImageFile:$WimFile" /Index:1 "/ApplyDir:$DriveLetter"
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -eq 0) {
            Write-Log "WinPE applied successfully." -Level Success

            # Create marker file on WinPE partition for safety
            $markerFile = Join-Path $DriveLetter "DEPLOYKEY.marker"
            $markerContent = @"

# USB DEPLOYMENT KEY MARKER FILE
# DO NOT DELETE - This file prevents accidental disk formatting
# Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Partition: WinPE
DEPLOYKEY_MARKER=TRUE
"@
            Set-Content -Path $markerFile -Value $markerContent -Force
            Write-Log "Created safety marker on WinPE partition." -Level Success

            # Inject up-to-date startnet.cmd from Scripts\ into boot.wim
            $startnetOk = Update-StartnetCmd -WinPEDriveLetter $DriveLetter
            if (-not $startnetOk) {
                Write-Log "Warning: startnet.cmd could not be updated in boot.wim." -Level Warning
            }

            return $true
        } else {
            Write-Log "DISM failed with exit code: $dismExitCode" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Error applying WinPE: $_" -Level Error
        return $false
    }
}

function Update-StartnetCmd {
    <#
    .SYNOPSIS
        Injects the current startnet.cmd into the WinPE boot.wim
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WinPEDriveLetter
    )

    $bootWim = Join-Path $WinPEDriveLetter "sources\boot.wim"
    $startnetSource = Join-Path $PSScriptRoot "Scripts\startnet.cmd"
    $mountDir = "C:\WinPE_Mount"

    if (-not (Test-Path $bootWim)) {
        Write-Log "boot.wim not found at: $bootWim" -Level Error
        return $false
    }

    if (-not (Test-Path $startnetSource)) {
        Write-Log "startnet.cmd not found at: $startnetSource" -Level Error
        return $false
    }

    # Clean up and create mount directory
    if (Test-Path $mountDir) {
        $p = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        Remove-Item -Path $mountDir -Recurse -Force -ErrorAction SilentlyContinue
        $ProgressPreference = $p
    }
    New-Item -Path $mountDir -ItemType Directory -Force | Out-Null

    # Proactive stale mount cleanup before attempting mount
    Remove-StaleWimMount -WimFilePath $bootWim | Out-Null

    try {
        # Mount boot.wim — use call operator to avoid Start-Process -Wait hanging
        # on DismHost.exe child processes that outlive dism.exe itself
        Write-Log "Mounting boot.wim to inject startnet.cmd..." -Level Info
        Write-Host ""
        & dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            Write-Log "Mount failed (exit code: $dismExitCode) — attempting stale mount recovery..." -Level Warning
            if (Remove-StaleWimMount -WimFilePath $bootWim) {
                Write-Log "Retrying mount..." -Level Info
                if (Test-Path $mountDir) { $p = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'; Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue; $ProgressPreference = $p }
                New-Item -Path $mountDir -ItemType Directory -Force | Out-Null
                Write-Host ""
                & dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
                $dismExitCode = $LASTEXITCODE
            }
        }

        if ($dismExitCode -ne 0) {
            Write-Log "Failed to mount boot.wim (exit code: $dismExitCode)" -Level Error
            return $false
        }

        # Copy startnet.cmd into the mounted image
        $startnetDest = Join-Path $mountDir "Windows\System32\startnet.cmd"
        Copy-Item -Path $startnetSource -Destination $startnetDest -Force
        Write-Log "startnet.cmd copied into WinPE image." -Level Success

        # Unmount and commit
        Write-Log "Committing boot.wim..." -Level Info
        Write-Host ""
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Commit
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            Write-Log "Failed to commit boot.wim (exit code: $dismExitCode)" -Level Error
            & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
            return $false
        }

        Write-Log "startnet.cmd updated in WinPE successfully." -Level Success
        return $true
    }
    catch {
        Write-Log "Error updating startnet.cmd: $_" -Level Error
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
        return $false
    }
    finally {
        if ((Test-Path $mountDir) -and ((Get-ChildItem $mountDir -Force | Measure-Object).Count -eq 0)) {
            Remove-Item $mountDir -Force -ErrorAction SilentlyContinue
        }
    }
}

function Add-WinPEDrivers {
    <#
    .SYNOPSIS
        Injects drivers into WinPE boot.wim
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WinPEDriveLetter,

        [Parameter(Mandatory)]
        [string]$DriverPath
    )

    Write-Section "Inject Drivers into WinPE"

    $bootWim = Join-Path $WinPEDriveLetter "sources\boot.wim"
    $mountDir = "C:\WinPE_Mount"

    # Validate boot.wim exists
    if (-not (Test-Path $bootWim)) {
        Write-Log "boot.wim not found at: $bootWim" -Level Error
        return $false
    }

    # Validate driver path exists
    if (-not (Test-Path $DriverPath -PathType Container)) {
        Write-Log "Driver path not found: $DriverPath" -Level Error
        return $false
    }

    # Always clean up stale mounts first — covers re-runs after Ctrl-C
    Remove-StaleWimMount -WimFilePath $bootWim | Out-Null

    # Clean and recreate mount directory
    if (Test-Path $mountDir) {
        $p = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        Remove-Item -Path $mountDir -Recurse -Force -ErrorAction SilentlyContinue
        $ProgressPreference = $p
    }
    New-Item -Path $mountDir -ItemType Directory -Force | Out-Null

    try {
        # Mount boot.wim — use call operator to avoid Start-Process -Wait hanging
        # on DismHost.exe child processes that outlive dism.exe itself
        Write-Log "Mounting boot.wim..." -Level Info
        Write-Host ""
        & dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            # Fallback stale-mount recovery (e.g. drive letter changed between sessions)
            Write-Log "Mount failed (exit code: $dismExitCode) — attempting stale mount recovery..." -Level Warning
            if (Remove-StaleWimMount -WimFilePath $bootWim) {
                Write-Log "Retrying mount..." -Level Info
                if (Test-Path $mountDir) { $p = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'; Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue; $ProgressPreference = $p }
                New-Item -Path $mountDir -ItemType Directory -Force | Out-Null
                Write-Host ""
                & dism.exe /Mount-Image "/ImageFile:$bootWim" /Index:1 "/MountDir:$mountDir"
                $dismExitCode = $LASTEXITCODE
            }
        }

        if ($dismExitCode -ne 0) {
            Write-Log "Failed to mount boot.wim (exit code: $dismExitCode)" -Level Error
            return $false
        }

        Write-Log "boot.wim mounted successfully." -Level Success

        # Inject drivers — RST/storage drivers can take several minutes, this is normal
        Write-Log "Injecting drivers from: $DriverPath" -Level Info
        Write-Log "(Driver injection may take several minutes — please wait)" -Level Warning
        Write-Host ""
        & dism.exe /Image:"$mountDir" /Add-Driver "/Driver:$DriverPath" /Recurse
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            Write-Log "Warning: Some drivers may have failed to inject (exit code: $dismExitCode)" -Level Warning
        } else {
            Write-Log "Drivers injected successfully." -Level Success
        }

        # Unmount and commit
        Write-Log "Unmounting and committing changes..." -Level Info
        Write-Host ""
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Commit
        $dismExitCode = $LASTEXITCODE

        if ($dismExitCode -ne 0) {
            Write-Log "Failed to unmount boot.wim (exit code: $dismExitCode)" -Level Error
            & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
            return $false
        }

        Write-Log "WinPE drivers injected and saved successfully." -Level Success
        return $true
    }
    catch {
        Write-Log "Error injecting drivers: $_" -Level Error
        & dism.exe /Unmount-Image "/MountDir:$mountDir" /Discard | Out-Null
        return $false
    }
    finally {
        if ((Test-Path $mountDir) -and ((Get-ChildItem $mountDir -Force | Measure-Object).Count -eq 0)) {
            Remove-Item $mountDir -Force -ErrorAction SilentlyContinue
        }
    }
}

function Copy-ScriptFiles {
    <#
    .SYNOPSIS
        Copies deployment scripts to the USB drive
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DriveLetter
    )

    Write-Section "Copy Script Files"

    $sourceDir = Join-Path $PSScriptRoot "Scripts"

    if (-not (Test-Path $sourceDir)) {
        Write-Log "Scripts directory not found: $sourceDir" -Level Error
        return $false
    }

    $destination = "${DriveLetter}\"

    Write-Log "Copying scripts to $DriveLetter..." -Level Info

    try {
        $files = Get-ChildItem -Path $sourceDir -File
        $totalFiles = $files.Count
        $counter = 0

        foreach ($file in $files) {
            $counter++
            Write-Progress -Activity "Copying Scripts" -Status $file.Name -PercentComplete (($counter / $totalFiles) * 100)

            # Don't overwrite config.ini if it already exists on the destination
            $destFile = Join-Path $destination $file.Name
            if ($file.Name -eq "config.ini" -and (Test-Path $destFile)) {
                Write-Log "  Skipping config.ini (already exists on destination)" -Level Info
                continue
            }

            Copy-Item -Path $file.FullName -Destination $destination -Force
        }

        Write-Progress -Activity "Copying Scripts" -Completed
        Write-Log "Copied $totalFiles script files." -Level Success

        # Create USB marker file for safety (prevents accidental formatting)
        $markerFile = Join-Path $destination "DEPLOYKEY.marker"
        $markerContent = @"
# USB DEPLOYMENT KEY MARKER FILE
# DO NOT DELETE - This file prevents accidental disk formatting
# Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# If this file exists on a disk, ApplyImage.bat will refuse to format it.
DEPLOYKEY_MARKER=TRUE
"@
        Set-Content -Path $markerFile -Value $markerContent -Force
        Write-Log "Created safety marker file: DEPLOYKEY.marker" -Level Success

        return $true
    }
    catch {
        Write-Log "Error copying scripts: $_" -Level Error
        return $false
    }
}

function Copy-ModelWIMs {
    <#
    .SYNOPSIS
        Copies model-specific WIM files (and optional drivers) to the USB drive.
        Supports two-level LNK redirect markers to avoid WIM duplication.

        Source directory structure under WimSourceDirectory:
          <ModelName>\                          - self-contained: all WIMs + optional Drivers
          <ModelName>\model.ini                 - overrides model string (ModelString=<value>)
          <ModelName>\LNK-<Src>.txt             - Level 1 redirect: if <Src> is a model dir, WIMs come from <Src>\
          <Src>\LNK-<Ver>.txt                   - Level 2 redirect: if <Ver> is not a model dir, OS from Windows_WIM_Root\<Ver>\
          <ModelName>\Drivers\LNK-<Src>.txt     - Driver redirect: drivers from Shared_Drivers_Location\<Src>\

        Selection label examples:
          B9450FA (Pure)                                no redirects, no Drivers
          B9450FA (Pure + Drivers)                      no redirects, own Drivers present
          B9450FA (Pure + Drivers->B5402CB)             drivers from B5402CB\Drivers\
          B5402CB (Pure)  ->  B9450FA                   Level 1 only: OS from B9450FA\
          B5402CB (Pure)  ->  B9450FA  [OS: W11-25H2]   Level 1+2: OS from WinWimRoot
          B5402CB (Pure)  [OS: W11-25H2]                Level 2 only: OS from WinWimRoot
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DriveLetter,

        [Parameter(Mandatory)]
        [string]$WimSourceDirectory,

        [string]$WinWimRoot,

        [string]$SharedDriversLocation
    )

    Write-Section "Copy Model WIMs"

    if (-not (Test-Path $WimSourceDirectory -PathType Container)) {
        Write-Log "WIM source directory not found: $WimSourceDirectory" -Level Error
        return $false
    }

    $modelDirs = Get-ChildItem -Path $WimSourceDirectory -Directory

    if ($modelDirs.Count -eq 0) {
        Write-Log "No directories found in: $WimSourceDirectory" -Level Warning
        return $false
    }

    # ============================================================
    # Build enriched entry list
    # ============================================================
    $entries = foreach ($dir in $modelDirs) {
        $hasModelIni = Test-Path (Join-Path $dir.FullName "model.ini")
        $hasDrivers  = Test-Path (Join-Path $dir.FullName "Drivers") -PathType Container

        # Detect driver LNK: LNK-<model>.txt inside the Drivers\ folder
        $lnkDrivers = $null
        if ($hasDrivers) {
            $drvLnkFile = Get-ChildItem -Path (Join-Path $dir.FullName "Drivers") -File `
                              -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match '^LNK-(.+)\.txt$' } |
                          Select-Object -First 1
            if ($drvLnkFile) {
                $lnkDrivers = [regex]::Match($drvLnkFile.Name, '^LNK-(.+)\.txt$').Groups[1].Value
            }
        }

        $hasAnyLnk = (Get-ChildItem -Path $dir.FullName -File -Filter "LNK-*.txt" `
                          -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0

        # Allow any directory whose name matches the "<ModelCode> - <desc>" convention
        $matchesConvention = $dir.Name -match '^[A-Za-z0-9]+ - '

        if (-not $hasModelIni -and -not $hasAnyLnk -and -not $lnkDrivers -and -not $matchesConvention -and $dir.Name.Length -gt 10) {
            Write-Log "Skipping '$($dir.Name)': name too long and no model.ini or LNK found." -Level Warning
            continue
        }

        # Resolve ModelString
        # Priority: model.ini > convention (part before first ' - ') > dir name
        $modelString = $dir.Name
        if ($hasModelIni) {
            $modelLine = Get-Content (Join-Path $dir.FullName "model.ini") -ErrorAction SilentlyContinue |
                         Where-Object { $_ -match '^ModelString=' } | Select-Object -First 1
            if ($modelLine) {
                $parsed = ($modelLine -split '=', 2)[1].Trim()
                if ($parsed) { $modelString = $parsed }
            }
        } elseif ($dir.Name -match '^([^ ]+) - ') {
            $modelString = $Matches[1]
        }

        # LNK detection: all LNK files use .txt extension.
        # A LNK whose target resolves to a model directory = Level 1 (WIM source redirect).
        # A LNK whose target does not resolve to a model directory = Level 2 (OS version).
        $lnkSource     = $null
        $lnkSourcePath = $null
        $lnkOs         = $null

        $rootLnk = Get-ChildItem -Path $dir.FullName -File -Filter "LNK-*.txt" `
                       -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($rootLnk) {
            $lnkTarget    = [regex]::Match($rootLnk.Name, '^LNK-(.+)\.txt$').Groups[1].Value
            $resolvedPath = Resolve-ModelDirectory -ModelName $lnkTarget -WimSourceDirectory $WimSourceDirectory
            if ($resolvedPath) {
                # Level 1: target is a model directory — WIMs come from there
                $lnkSource     = $lnkTarget
                $lnkSourcePath = $resolvedPath
                # Level 2: look for OS version LNK inside the Level 1 source dir
                $lnkOsFile = Get-ChildItem -Path $lnkSourcePath -File -Filter "LNK-*.txt" `
                                 -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($lnkOsFile) {
                    $lnkOs = [regex]::Match($lnkOsFile.Name, '^LNK-(.+)\.txt$').Groups[1].Value
                }
            } else {
                # Level 2 only: no Level 1 redirect, target is an OS version name
                $lnkOs = $lnkTarget
            }
        }

        # Build display label
        $label = if (-not $hasModelIni) {
            $drvTag = if (-not $hasDrivers)  { ' (Pure)' }
                      elseif ($lnkDrivers)   { " (Pure + Drivers->$lnkDrivers)" }
                      else                   { ' (Pure + Drivers)' }
            "$($dir.Name)$drvTag"
        } else {
            $drvTag = if (-not $hasDrivers)  { '' }
                      elseif ($lnkDrivers)   { " (Drivers->$lnkDrivers)" }
                      else                   { ' (Drivers)' }
            "$($dir.Name) [$modelString]$drvTag"
        }

        if ($lnkSource -and $lnkOs) {
            $label += "  ->  $lnkSource  [OS: $lnkOs]"
        } elseif ($lnkSource) {
            $label += "  ->  $lnkSource"
        } elseif ($lnkOs) {
            $label += "  [OS: $lnkOs]"
        }

        [PSCustomObject]@{
            Label          = $label
            DirName        = $dir.Name
            Path           = $dir.FullName
            HasModelIni    = $hasModelIni
            HasDrivers     = $hasDrivers
            ModelString    = $modelString
            LnkSource      = $lnkSource
            LnkSourcePath  = $lnkSourcePath    # resolved directory path
            LnkOs          = $lnkOs
            LnkDrivers     = $lnkDrivers
        }
    }

    Write-Host ""
    Write-Host "  Available Models:" -ForegroundColor Cyan
    $selectedLabel = Select-FromList -Items ($entries | Select-Object -ExpandProperty Label) -Prompt "Select model"
    $selected = $entries | Where-Object { $_.Label -eq $selectedLabel }

    $modelString = $selected.ModelString
    Write-Log "Selected: $($selected.DirName)  ->  ModelString: $modelString" -Level Info

    $destination = "${DriveLetter}\"

    # ============================================================
    # Resolve WIM file list with source paths and destination names
    # ============================================================
    # Each entry: [PSCustomObject]@{ SourcePath; DestName; Length }
    $filesToCopy = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($selected.LnkSource) {
        # Level 1 redirect: OS (and un-overridden WIMs) come from the source model dir.
        # SYSTEM / RECOVERY / MYASUS WIMs may be present directly in the model dir.
        $sourceDir = $selected.LnkSourcePath

        if (-not $sourceDir -or -not (Test-Path $sourceDir -PathType Container)) {
            Write-Log "Level 1 redirect target not found for: $($selected.LnkSource)" -Level Error
            Write-Log "  Expected a directory under: $WimSourceDirectory" -Level Warning
            return $false
        }

        # WIMs already present in the model dir — rename to targetModel convention
        foreach ($f in (Get-ChildItem -Path $selected.Path -File |
                         Where-Object { $_.Extension -in @('.wim', '.swm') })) {
            $destName = Get-WimDestName -SourceName $f.Name -TargetModel $modelString
            $filesToCopy.Add([PSCustomObject]@{ SourcePath = $f.FullName; DestName = $destName; Length = $f.Length })
        }

        # WIMs from Level 1 source dir — rename to targetModel convention
        $sourceDirWims = @(Get-ChildItem -Path $sourceDir -File |
                           Where-Object { $_.Extension -in @('.wim', '.swm') })

        if ($selected.LnkOs) {
            # Level 2: exclude OS files from the source dir (they come from WinWimRoot)
            $sourceDirWims = @($sourceDirWims | Where-Object {
                $basePart = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $partType = if ($basePart -match '-([^-]+)$') { $Matches[1] } else { $basePart }
                $partType -notmatch '^OS'
            })
        }

        foreach ($f in $sourceDirWims) {
            $destName = Get-WimDestName -SourceName $f.Name -TargetModel $modelString
            $filesToCopy.Add([PSCustomObject]@{ SourcePath = $f.FullName; DestName = $destName; Length = $f.Length })
        }

        # Level 2: OS from WinWimRoot
        if ($selected.LnkOs) {
            if ([string]::IsNullOrWhiteSpace($WinWimRoot)) {
                Write-Log "LNK-OS redirect requires Windows_WIM_Root in config.psd1" -Level Error
                return $false
            }
            $externalOsPath = Join-Path $WinWimRoot "$($selected.LnkOs)\$($selected.LnkOs).wim"
            if (-not (Test-Path $externalOsPath)) {
                Write-Log "LNK-OS-$($selected.LnkOs): file not found: $externalOsPath" -Level Error
                return $false
            }
            $osFile = Get-Item $externalOsPath
            Write-Log "OS WIM source: $externalOsPath  [LNK-OS-$($selected.LnkOs)]" -Level Info
            $filesToCopy.Add([PSCustomObject]@{ SourcePath = $osFile.FullName; DestName = "$modelString-OS.wim"; Length = $osFile.Length })
        }

    } elseif ($selected.LnkOs) {
        # Level 2 only (no Level 1): OS from WinWimRoot, other WIMs from model dir
        if ([string]::IsNullOrWhiteSpace($WinWimRoot)) {
            Write-Log "LNK-OS redirect requires Windows_WIM_Root in config.psd1" -Level Error
            return $false
        }
        $externalOsPath = Join-Path $WinWimRoot "$($selected.LnkOs)\$($selected.LnkOs).wim"
        if (-not (Test-Path $externalOsPath)) {
            Write-Log "LNK-OS-$($selected.LnkOs): file not found: $externalOsPath" -Level Error
            return $false
        }
        $osFile = Get-Item $externalOsPath
        Write-Log "OS WIM source: $externalOsPath  [LNK-OS-$($selected.LnkOs)]" -Level Info
        $filesToCopy.Add([PSCustomObject]@{ SourcePath = $osFile.FullName; DestName = "$modelString-OS.wim"; Length = $osFile.Length })

        # Other WIMs from model dir (exclude OS) — rename to targetModel convention
        foreach ($f in (Get-ChildItem -Path $selected.Path -File |
                         Where-Object { $_.Extension -in @('.wim', '.swm') })) {
            $basePart = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            $partType = if ($basePart -match '-([^-]+)$') { $Matches[1] } else { $basePart }
            if ($partType -match '^OS') { continue }
            $destName = Get-WimDestName -SourceName $f.Name -TargetModel $modelString
            $filesToCopy.Add([PSCustomObject]@{ SourcePath = $f.FullName; DestName = $destName; Length = $f.Length })
        }

    } else {
        # No LNK: copy all WIMs from model dir — rename to targetModel convention
        foreach ($f in (Get-ChildItem -Path $selected.Path -File |
                         Where-Object { $_.Extension -in @('.wim', '.swm') })) {
            $destName = Get-WimDestName -SourceName $f.Name -TargetModel $modelString
            $filesToCopy.Add([PSCustomObject]@{ SourcePath = $f.FullName; DestName = $destName; Length = $f.Length })
        }
    }

    if ($filesToCopy.Count -eq 0) {
        Write-Log "No WIM/SWM files resolved to copy." -Level Warning
        return $false
    }

    # ============================================================
    # Update config.ini (ModelString + EditionIndex)
    # ============================================================
    $configFile = Join-Path $destination "config.ini"
    if (Test-Path $configFile) {
        $content = Get-Content $configFile
        $content = $content -replace '^ModelString=.*', "ModelString=$modelString"
        Write-Log "Updating config.ini: ModelString=$modelString" -Level Info

        # Find OS WIM source for Pro edition detection
        $osEntry = $filesToCopy | Where-Object { $_.DestName -match '-OS\.wim$' } | Select-Object -First 1
        if ($osEntry) {
            try {
                $images = Get-WindowsImage -ImagePath $osEntry.SourcePath -ErrorAction Stop
                $proImage = $images | Where-Object { $_.ImageName -match '\bPro\b' } | Select-Object -First 1

                if ($proImage) {
                    $detectedIndex = $proImage.ImageIndex
                    Write-Log "Detected Pro edition: '$($proImage.ImageName)' at index $detectedIndex" -Level Success
                    if ($content -match '^EditionIndex=') {
                        $content = $content -replace '^EditionIndex=.*', "EditionIndex=$detectedIndex"
                    } else {
                        $content += "`nEditionIndex=$detectedIndex"
                    }
                    Write-Log "Updating config.ini: EditionIndex=$detectedIndex" -Level Info
                } elseif ($images.Count -eq 1) {
                    Write-Log "Single-edition WIM ($($images[0].ImageName)) — EditionIndex left as-is." -Level Info
                } else {
                    Write-Log "No Pro edition found in WIM — EditionIndex left as-is." -Level Warning
                    Write-Log "Available: $(($images | Select-Object -ExpandProperty ImageName) -join ', ')" -Level Warning
                }
            }
            catch {
                Write-Log "Could not read WIM image info: $_" -Level Warning
                Write-Log "EditionIndex left as-is in config.ini." -Level Warning
            }
        } else {
            Write-Log "No OS WIM resolved — EditionIndex left as-is." -Level Warning
        }

        Set-Content -Path $configFile -Value $content
    }

    # ============================================================
    # Copy WIM files
    # ============================================================
    $totalSizeBytes = ($filesToCopy | Measure-Object -Property Length -Sum).Sum
    $totalSizeMB    = [math]::Round($totalSizeBytes / 1MB, 0)
    Write-Log "Copying $($filesToCopy.Count) WIM/SWM file(s) (${totalSizeMB} MB total)..." -Level Info

    $copiedBytes = 0L
    foreach ($entry in $filesToCopy) {
        $destPath = Join-Path $destination $entry.DestName
        $srcName  = [System.IO.Path]::GetFileName($entry.SourcePath)
        if ($srcName -ne $entry.DestName) {
            Write-Log "  $srcName  ->  $($entry.DestName)" -Level Info
        }
        Copy-FileWithProgress -Source $entry.SourcePath -Destination $destPath `
                              -TotalBatchBytes $totalSizeBytes -BytesBefore $copiedBytes `
                              -Activity "Copying WIM Files"
        $copiedBytes += $entry.Length
    }

    Write-Progress -Activity "Copying WIM Files" -Completed

    $totalSizeGB = [math]::Round($totalSizeBytes / 1GB, 2)
    Write-Log "Copied $($filesToCopy.Count) WIM/SWM file(s) (${totalSizeGB} GB total)." -Level Success

    # ============================================================
    # Copy drivers — resolves LNK-<model>.txt redirect in Drivers\ if present
    # ============================================================
    if ($selected.HasDrivers) {
        $driversDir     = Join-Path $selected.Path "Drivers"
        $driverResolved = Resolve-DriverLnk -DriversDir $driversDir -SharedDriversLocation $SharedDriversLocation

        if ($driverResolved) {
            if ($driverResolved.LinkedModel) {
                Write-Log "Drivers: following LNK to $($driverResolved.LinkedModel)\Drivers..." -Level Info
            }

            $driverSource = $driverResolved.Path
            $driverDest   = Join-Path $destination "Drivers\$modelString"

            Write-Log "Copying drivers to Drivers\$modelString..." -Level Info

            try {
                New-Item -Path $driverDest -ItemType Directory -Force | Out-Null
                Copy-Item -Path "$driverSource\*" -Destination $driverDest -Recurse -Force

                $driverCount = (Get-ChildItem -Path $driverDest -Recurse -File).Count
                Write-Log "Drivers copied ($driverCount files -> Drivers\$modelString)." -Level Success
            }
            catch {
                Write-Log "Error copying drivers: $_" -Level Warning
            }
        }
    }

    return $true
}

# ============================================================
# Main Script
# ============================================================

# Initialize logging
Write-Banner "ProduceKey - USB Key Creation Tool"
if ($Script:LogEnabled) {
    Write-Log "Log file: $Script:LogFile" -Level Info
}
Write-Log "Started at: $(Get-Date)" -Level Info

# Load configuration
Write-Section "Load Configuration"

$configPath = Join-Path $PSScriptRoot "config.psd1"

if (-not (Test-Path $configPath)) {
    Write-Log "Configuration file not found: $configPath" -Level Error
    exit 1
}

$config       = Import-PowerShellDataFile -Path $configPath
$PE_WIM_ROOT      = $config["WinPE_WIM_Location"]
$DEV_ROOT_WIM     = $config["Device_Root_WIM_Location"]
$WIN_WIM_ROOT     = $config["Windows_WIM_Root"]
$SHARED_DRV_ROOT  = $config["Shared_Drivers_Location"]

Write-Log "WinPE root: $PE_WIM_ROOT" -Level Info
Write-Log "Device WIM Root: $DEV_ROOT_WIM" -Level Info
if ($WIN_WIM_ROOT)    { Write-Log "Windows WIM Root: $WIN_WIM_ROOT" -Level Info }
if ($SHARED_DRV_ROOT) { Write-Log "Shared Drivers: $SHARED_DRV_ROOT" -Level Info }

# ============================================================
# Select WinPE version
# ============================================================
Write-Section "Select WinPE Version"

$PE_WIM     = $null
$peWimValid = $false

if (-not (Test-Path $PE_WIM_ROOT -PathType Container)) {
    Write-Log "WinPE directory not found: $PE_WIM_ROOT" -Level Warning
    Write-Log "Run ExtractPE.ps1 to populate this directory." -Level Warning
} else {
    $peEntries = @(Get-ChildItem -Path $PE_WIM_ROOT -Directory |
                   Where-Object { Test-Path (Join-Path $_.FullName "boot.wim") } |
                   ForEach-Object { [PSCustomObject]@{ Label = $_.Name; WimPath = Join-Path $_.FullName "boot.wim" } })

    if ($peEntries.Count -eq 0) {
        Write-Log "No boot.wim found in any subdirectory of: $PE_WIM_ROOT" -Level Warning
        Write-Log "Run ExtractPE.ps1 to populate this directory." -Level Warning
    } else {
        $selectedLabel = Select-FromList -Items ($peEntries | Select-Object -ExpandProperty Label) `
                                         -Prompt "Select WinPE version to use"
        $selectedPE    = $peEntries | Where-Object { $_.Label -eq $selectedLabel }
        $PE_WIM        = $selectedPE.WimPath
        $peWimValid    = $true
        Write-Log "Using WinPE: $PE_WIM" -Level Success
    }
}

# Track what drive letters we're using
$winPEDrive = $null
$imagesDrive = $null

# ============================================================
# Step 1: Format USB Drive (Optional)
# ============================================================
Write-Section "Step 1: Format USB Drive"

if (Read-YesNo "Do you want to format a USB drive?") {

    if (-not $peWimValid) {
        Write-Log "Cannot proceed: WinPE WIM file is required for formatting." -Level Error
    } else {
        $formatResult = Format-USBDrive

        if ($formatResult) {
            $winPEDrive = $formatResult.WinPELetter
            $imagesDrive = $formatResult.ImagesLetter

            # Automatically install WinPE after format
            Write-Host ""
            if (Read-YesNo "Apply WinPE to the new partition?" -Default $true) {
                Install-WinPE -WimFile $PE_WIM -DriveLetter $winPEDrive
            }
        }
    }
} else {
    # ============================================================
    # Step 2: Apply WinPE Only (Optional)
    # ============================================================
    Write-Section "Step 2: Apply WinPE (Optional)"

    if ($peWimValid -and (Read-YesNo "Do you want to apply WinPE to an existing partition?")) {
        $selectedDrive = Select-DriveLetter -Purpose "WinPE installation"

        if ($selectedDrive) {
            $winPEDrive = $selectedDrive.DriveLetter
            Install-WinPE -WimFile $PE_WIM -DriveLetter $winPEDrive
        }
    }
}

# ============================================================
# Step 3: Copy Script Files (Optional)
# ============================================================
Write-Section "Step 3: Copy Script Files (Optional)"

if (Read-YesNo "Do you want to copy deployment scripts to USB?") {
    Write-Host ""
    Write-Log "Select the destination drive for deployment scripts." -Level Info

    $scriptDrive = Select-DriveLetter -Purpose "script files"

    if ($scriptDrive) {
        $imagesDrive = $scriptDrive.DriveLetter
        Copy-ScriptFiles -DriveLetter $imagesDrive
    } else {
        Write-Log "No drive selected for scripts. Skipping." -Level Warning
    }
}

# ============================================================
# Step 4: Copy Model WIMs (Optional)
# ============================================================
Write-Section "Step 4: Copy Model WIMs (Optional)"

if (Read-YesNo "Do you want to copy model WIM files to USB?") {

    # Get Images drive if not already known from a previous step
    if (-not $imagesDrive) {
        $selectedDrive = Select-DriveLetter -Purpose "Images partition (NTFS)"
        if ($selectedDrive) { $imagesDrive = $selectedDrive.DriveLetter }
    }

    if (-not $imagesDrive) {
        Write-Log "No drive selected. Skipping." -Level Warning
    } elseif (-not (Test-Path -Path $DEV_ROOT_WIM -PathType Container)) {
        Write-Log "Device WIM root directory not found: $DEV_ROOT_WIM" -Level Error
        Write-Log "Please update Device_Root_WIM_Location in config.psd1" -Level Warning
    } else {
        Copy-ModelWIMs -DriveLetter $imagesDrive -WimSourceDirectory $DEV_ROOT_WIM -WinWimRoot $WIN_WIM_ROOT -SharedDriversLocation $SHARED_DRV_ROOT
    }
}

# ============================================================
# Step 5: Inject Drivers into WinPE (Optional)
# ============================================================
Write-Section "Step 5: Inject Drivers into WinPE (Optional)"

Write-Host ""
Write-Log "Inject storage drivers (e.g., Intel RST) into WinPE to enable disk detection." -Level Info
Write-Host ""

if (Read-YesNo "Do you want to inject drivers into WinPE boot.wim?") {

    # Get WinPE partition if not already known
    if (-not $winPEDrive) {
        Write-Host ""
        Write-Log "Select the WinPE partition (FAT32 partition with sources\boot.wim):" -Level Info
        $selectedPE = Select-DriveLetter -Purpose "WinPE partition"
        if ($selectedPE) { $winPEDrive = $selectedPE.DriveLetter }
    }

    if (-not $winPEDrive) {
        Write-Log "No WinPE partition selected. Skipping." -Level Warning
    } else {
        $bootWimPath = Join-Path $winPEDrive "sources\boot.wim"
        if (-not (Test-Path $bootWimPath)) {
            Write-Log "boot.wim not found at $bootWimPath - is this the correct WinPE partition?" -Level Error
        } else {

            # --- Discover RST driver sources ---
            $rstEntries = @()

            # 1. Check drivers already on the images drive (copied by Step 4)
            if ($imagesDrive) {
                $driversRoot = Join-Path $imagesDrive "Drivers"
                if (Test-Path $driversRoot -PathType Container) {
                    foreach ($modelDir in (Get-ChildItem $driversRoot -Directory)) {
                        $rst = Find-RSTDriverPath -DriverDir $modelDir.FullName
                        if ($rst) {
                            $rstEntries += [PSCustomObject]@{
                                Label = "$($modelDir.Name)  [on key: $imagesDrive]"
                                Path  = $rst
                            }
                        }
                    }
                }
            }

            # 2. Nothing on the key — scan the WIM repository for models with Drivers\
            if ($rstEntries.Count -eq 0) {
                if (Test-Path $DEV_ROOT_WIM -PathType Container) {
                    Write-Log "No drivers found on key — scanning WIM repository for driver sources..." -Level Info
                    foreach ($dir in (Get-ChildItem $DEV_ROOT_WIM -Directory)) {
                        $srcDriverDir = Join-Path $dir.FullName "Drivers"
                        if (Test-Path $srcDriverDir -PathType Container) {
                            $driverResolved = Resolve-DriverLnk -DriversDir $srcDriverDir -SharedDriversLocation $SHARED_DRV_ROOT
                            if ($driverResolved) {
                                $rst = Find-RSTDriverPath -DriverDir $driverResolved.Path
                                if ($rst) {
                                    $lnkSuffix = if ($driverResolved.LinkedModel) { "  [->$($driverResolved.LinkedModel)]" } else { '' }
                                    $rstEntries += [PSCustomObject]@{
                                        Label = "$($dir.Name)$lnkSuffix  [from repository]"
                                        Path  = $rst
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if ($rstEntries.Count -eq 0) {
                Write-Log "No RST/storage driver sources found. Skipping." -Level Warning
                Write-Log "Expected structure: <Images>:\Drivers\<Model>\Rapid Storage\  or  <WIM repo>\<Model>\Drivers\Rapid Storage\" -Level Warning
            } else {
                # Always go through the list so the user can confirm before injecting
                Write-Host ""
                $driverMsg = if ($rstEntries.Count -eq 1) { 'Driver source found' } else { 'Multiple driver sources found' }
                Write-Log "$driverMsg — select one to inject:" -Level Info
                $selectedLabel = Select-FromList -Items ($rstEntries | Select-Object -ExpandProperty Label) `
                                                -Prompt "Select driver source"
                $selectedEntry = $rstEntries | Where-Object { $_.Label -eq $selectedLabel }

                if ($selectedEntry) {
                    Write-Log "Using driver path: $($selectedEntry.Path)" -Level Info
                    Add-WinPEDrivers -WinPEDriveLetter $winPEDrive -DriverPath $selectedEntry.Path
                }
            }
        }
    }
}

# ============================================================
# Summary
# ============================================================
Write-Banner "Complete"

Write-Log "USB key preparation finished." -Level Success
if ($Script:LogEnabled) {
    Write-Log "Log saved to: $Script:LogFile" -Level Info
}

if ($imagesDrive) {
    Write-Host ""
    Write-Host "  Next Steps:" -ForegroundColor Cyan
    Write-Host "  1. Check config.ini on $imagesDrive for ModelString and DiskNumber" -ForegroundColor White
    Write-Host "  2. Boot target device from USB key" -ForegroundColor White
    Write-Host "  3. Check log files on USB after deployment" -ForegroundColor White
    Write-Host ""
}

Write-Host "Merci d'avoir utilise ProduceKey!" -ForegroundColor Green
Write-Host ""
