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
    Requires: Administrator privileges, Windows ADK
#>
param(
    [switch]$Log
)

# ============================================================
# Configuration
# ============================================================
$Script:SizeLimit = 128GB  # Maximum drive size to allow formatting
$Script:LogEnabled = $Log.IsPresent
$Script:LogFile = if ($Script:LogEnabled) { "ProduceKey_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" } else { $null }

# ============================================================
# Logging Functions
# ============================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"

    # Write to log file (if enabled)
    if ($Script:LogEnabled) {
        Add-Content -Path $Script:LogFile -Value $logMessage
    }

    # Write to console with color
    switch ($Level) {
        'Info'    { Write-Host $Message }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error'   { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
    }
}

function Write-Banner {
    param([string]$Title)

    $line = "=" * 60
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""

    if ($Script:LogEnabled) {
        Add-Content -Path $Script:LogFile -Value ""
        Add-Content -Path $Script:LogFile -Value $line
        Add-Content -Path $Script:LogFile -Value "  $Title"
        Add-Content -Path $Script:LogFile -Value $line
    }
}

function Write-Section {
    param([string]$Title)

    Write-Host ""
    Write-Host "--- $Title ---" -ForegroundColor Magenta
    if ($Script:LogEnabled) {
        Add-Content -Path $Script:LogFile -Value ""
        Add-Content -Path $Script:LogFile -Value "--- $Title ---"
    }
}

# ============================================================
# User Interaction Functions
# ============================================================
function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $false
    )

    $defaultHint = if ($Default) { "[Y/n]" } else { "[y/N]" }
    Write-Host "$Prompt $defaultHint : " -NoNewline -ForegroundColor Yellow
    $response = Read-Host

    if ([string]::IsNullOrWhiteSpace($response)) {
        return $Default
    }

    return $response -match '^[Yy]'
}

function Read-KeyPress {
    param([string]$Prompt)

    Write-Host $Prompt -NoNewline -ForegroundColor Yellow
    $key = [System.Console]::ReadKey($true)
    Write-Host $key.KeyChar

    if ($key.Key -eq [ConsoleKey]::Escape) {
        return $null
    }

    return $key.KeyChar.ToString().ToUpper()
}

function Select-FromList {
    param(
        [Parameter(Mandatory)]
        [object[]]$Items,
        [string]$Prompt = "Please choose an option"
    )

    Write-Host ""
    for ($i = 0; $i -lt $Items.Length; $i++) {
        Write-Host "  [$($i + 1)] " -NoNewline -ForegroundColor Cyan
        Write-Host $Items[$i]
    }
    Write-Host ""

    do {
        Write-Host "$Prompt (1-$($Items.Length)): " -NoNewline -ForegroundColor Yellow
        $choice = Read-Host
        $selectedNumber = $choice -as [int]
    } while ($selectedNumber -lt 1 -or $selectedNumber -gt $Items.Length)

    return $Items[$selectedNumber - 1]
}

# ============================================================
# Drive Detection Functions
# ============================================================
function Get-USBDrives {
    <#
    .SYNOPSIS
        Lists all USB storage devices with their partitions
    #>

    $usbDisks = Get-Disk | Where-Object { $_.BusType -eq 'USB' -and $_.Size -gt 0 }

    if ($usbDisks.Count -eq 0) {
        Write-Log "No USB storage devices with media found." -Level Warning
        return @()
    }

    $drives = @()

    foreach ($disk in $usbDisks) {
        $partitions = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue

        if ($partitions) {
            foreach ($partition in $partitions) {
                $volume = Get-Volume -Partition $partition -ErrorAction SilentlyContinue
                $sizeGB = [math]::Round($partition.Size / 1GB, 2)
                $driveLetter = if ($volume.DriveLetter) { "$($volume.DriveLetter):" } else { $null }
                $fileSystem = if ($volume.FileSystem) { $volume.FileSystem } else { "RAW" }
                $label = if ($volume.FileSystemLabel) { $volume.FileSystemLabel } else { "No Label" }

                $drives += [PSCustomObject]@{
                    DiskNumber      = $disk.Number
                    DiskName        = $disk.FriendlyName
                    PartitionNumber = $partition.PartitionNumber
                    DriveLetter     = $driveLetter
                    SizeGB          = $sizeGB
                    FileSystem      = $fileSystem
                    Label           = $label
                    TotalDiskSizeGB = [math]::Round($disk.Size / 1GB, 2)
                }
            }
        } else {
            # Disk has no partitions
            $drives += [PSCustomObject]@{
                DiskNumber      = $disk.Number
                DiskName        = $disk.FriendlyName
                PartitionNumber = 0
                DriveLetter     = $null
                SizeGB          = 0
                FileSystem      = "Unallocated"
                Label           = "No Partitions"
                TotalDiskSizeGB = [math]::Round($disk.Size / 1GB, 2)
            }
        }
    }

    return $drives
}

function Show-USBDrives {
    <#
    .SYNOPSIS
        Displays USB drives in a formatted table
    #>
    param([array]$Drives)

    if ($Drives.Count -eq 0) {
        Write-Log "No USB drives found." -Level Warning
        return
    }

    Write-Host ""
    Write-Host "  USB Storage Devices Found:" -ForegroundColor Cyan
    Write-Host "  " + ("-" * 56) -ForegroundColor DarkGray

    foreach ($drive in $Drives) {
        $letterDisplay = if ($drive.DriveLetter) { $drive.DriveLetter } else { "--" }
        $info = "  Disk $($drive.DiskNumber) | $letterDisplay | $($drive.SizeGB)GB | $($drive.FileSystem) | $($drive.Label)"

        if ($drive.DriveLetter) {
            Write-Host $info -ForegroundColor White
        } else {
            Write-Host $info -ForegroundColor DarkGray
        }
    }

    Write-Host "  " + ("-" * 56) -ForegroundColor DarkGray
    Write-Host "  Device: " -NoNewline -ForegroundColor DarkGray

    $diskNames = $Drives | Select-Object -ExpandProperty DiskName -Unique
    Write-Host ($diskNames -join ", ") -ForegroundColor DarkGray
    Write-Host ""
}

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

function Select-DriveLetter {
    <#
    .SYNOPSIS
        Prompts user to select a drive letter from available USB partitions
    #>
    param(
        [string]$Purpose = "operation",
        [switch]$AllowUnpartitioned
    )

    $drives = Get-USBDrives
    Show-USBDrives -Drives $drives

    # Filter to drives with letters (and partitions)
    $availableDrives = $drives | Where-Object { $_.DriveLetter -ne $null }

    if ($availableDrives.Count -eq 0) {
        if ($AllowUnpartitioned -and ($drives | Where-Object { $_.FileSystem -eq 'Unallocated' })) {
            Write-Log "USB disk found but has no partitions. Format required." -Level Warning
            return $null
        }
        Write-Log "No USB partitions with drive letters found." -Level Error
        return $null
    }

    while ($true) {
        $letter = Read-KeyPress "Enter drive letter for $Purpose (or Esc to cancel): "

        if ($null -eq $letter) {
            Write-Log "Operation cancelled." -Level Warning
            return $null
        }

        $driveLetter = "${letter}:"
        $selectedDrive = $availableDrives | Where-Object { $_.DriveLetter -eq $driveLetter }

        if (-not $selectedDrive) {
            Write-Log "Drive $driveLetter is not a valid USB partition." -Level Warning
            continue
        }

        # Safety check - don't format drives with Windows directory
        if (Test-Path "$driveLetter\Windows") {
            Write-Log "Drive $driveLetter contains a Windows directory. Skipping for safety." -Level Warning
            continue
        }

        Write-Log "Selected: $driveLetter ($($selectedDrive.Label), $($selectedDrive.SizeGB)GB, $($selectedDrive.FileSystem))" -Level Success

        return [PSCustomObject]@{
            DriveLetter = $driveLetter
            DiskNumber  = $selectedDrive.DiskNumber
            SizeGB      = $selectedDrive.SizeGB
            FileSystem  = $selectedDrive.FileSystem
            Label       = $selectedDrive.Label
        }
    }
}

# ============================================================
# Core Operations
# ============================================================
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
create partition primary size=2048
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
        Write-Log "  WinPE partition: P: (2GB FAT32)" -Level Info
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
        Applies WinPE WIM to a partition
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
        $dismArgs = "/Apply-Image /ImageFile:`"$WimFile`" /Index:1 /ApplyDir:$DriveLetter /Verify"
        Write-Log "Running DISM..." -Level Info

        $process = Start-Process -FilePath "dism.exe" -ArgumentList $dismArgs -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0) {
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
            Write-Log "DISM failed with exit code: $($process.ExitCode)" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Error applying WinPE: $_" -Level Error
        return $false
    }
}

function Remove-StaleWimMount {
    <#
    .SYNOPSIS
        Finds and discards any stale DISM mount for a given WIM file.
        Uses Get-WindowsImage -Mounted (structured, non-localized) instead of
        parsing DISM text output so it works regardless of Windows language.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WimFilePath
    )

    Write-Log "Querying DISM for stale mounts of: $WimFilePath" -Level Warning

    try {
        $stale = Get-WindowsImage -Mounted -ErrorAction Stop |
                 Where-Object { $_.ImagePath -ieq $WimFilePath }

        if (-not $stale) {
            Write-Log "No stale DISM mount found for this WIM." -Level Warning
            return $false
        }

        $cleaned = $false
        foreach ($img in $stale) {
            Write-Log "Discarding stale mount at: $($img.MountPath)" -Level Warning
            $p = Start-Process -FilePath "dism.exe" `
                -ArgumentList "/Unmount-Image /MountDir:`"$($img.MountPath)`" /Discard" `
                -NoNewWindow -Wait -PassThru

            if ($p.ExitCode -eq 0) {
                Write-Log "Stale mount discarded." -Level Success
                $cleaned = $true
            } else {
                # Discard can fail if the WIM file was replaced (e.g. USB reformatted).
                # Still force a Cleanup-Wim so the orphaned record is cleared for the retry.
                Write-Log "Failed to discard stale mount (exit code: $($p.ExitCode)) - WIM file may have been replaced. Forcing Cleanup-Wim." -Level Warning
                $cleaned = $true
            }
        }

        if ($cleaned) {
            # Clean up DISM's internal mount metadata to prevent the retry from hanging
            Write-Log "Running Dism /Cleanup-Wim to clear residual mount metadata..." -Level Info
            Start-Process -FilePath "dism.exe" -ArgumentList "/Cleanup-Wim" -NoNewWindow -Wait | Out-Null
            Write-Log "DISM cleanup complete." -Level Info
        }

        return $cleaned
    }
    catch {
        Write-Log "Error querying mounted images: $_" -Level Error
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
        Remove-Item -Path $mountDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -Path $mountDir -ItemType Directory -Force | Out-Null

    try {
        # Mount boot.wim
        Write-Log "Mounting boot.wim to inject startnet.cmd..." -Level Info
        $process = Start-Process -FilePath "dism.exe" `
            -ArgumentList "/Mount-Image /ImageFile:`"$bootWim`" /Index:1 /MountDir:`"$mountDir`"" `
            -NoNewWindow -Wait -PassThru

        # 0xC1420127: image already mounted from a previous interrupted session
        if ($process.ExitCode -eq -1052638937) {
            Write-Log "boot.wim is already mounted (previous session was interrupted)." -Level Warning

            if (Remove-StaleWimMount -WimFilePath $bootWim) {
                Write-Log "Retrying mount..." -Level Info
                # Recreate a clean mount directory after the discard
                if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item -Path $mountDir -ItemType Directory -Force | Out-Null

                $process = Start-Process -FilePath "dism.exe" `
                    -ArgumentList "/Mount-Image /ImageFile:`"$bootWim`" /Index:1 /MountDir:`"$mountDir`"" `
                    -NoNewWindow -Wait -PassThru
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Failed to mount boot.wim (exit code: $($process.ExitCode))" -Level Error
            return $false
        }

        # Copy startnet.cmd into the mounted image
        $startnetDest = Join-Path $mountDir "Windows\System32\startnet.cmd"
        Copy-Item -Path $startnetSource -Destination $startnetDest -Force
        Write-Log "startnet.cmd copied into WinPE image." -Level Success

        # Unmount and commit
        Write-Log "Committing boot.wim..." -Level Info
        $process = Start-Process -FilePath "dism.exe" `
            -ArgumentList "/Unmount-Image /MountDir:`"$mountDir`" /Commit" `
            -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -ne 0) {
            Write-Log "Failed to commit boot.wim (exit code: $($process.ExitCode))" -Level Error
            Start-Process -FilePath "dism.exe" -ArgumentList "/Unmount-Image /MountDir:`"$mountDir`" /Discard" -NoNewWindow -Wait | Out-Null
            return $false
        }

        Write-Log "startnet.cmd updated in WinPE successfully." -Level Success
        return $true
    }
    catch {
        Write-Log "Error updating startnet.cmd: $_" -Level Error
        Start-Process -FilePath "dism.exe" -ArgumentList "/Unmount-Image /MountDir:`"$mountDir`" /Discard" -NoNewWindow -Wait | Out-Null
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
        Remove-Item -Path $mountDir -Recurse -Force -ErrorAction SilentlyContinue
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
                if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
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

function Find-RSTDriverPath {
    <#
    .SYNOPSIS
        Returns the RST/storage driver subfolder within a driver directory,
        or $null if none is found.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DriverDir
    )

    $rst = Join-Path $DriverDir "Rapid Storage"
    if (Test-Path $rst -PathType Container) { return $rst }

    $alt = Get-ChildItem -Path $DriverDir -Directory -Recurse -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -match 'Storage|RST|IRST' } |
           Select-Object -First 1

    return if ($alt) { $alt.FullName } else { $null }
}

function Copy-FileWithProgress {
    <#
    .SYNOPSIS
        Copies a single file with byte-level progress reporting.
        Designed for large WIM files where Copy-Item gives no feedback.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination,

        # Total bytes in the batch (for overall % across multiple files)
        [long]$TotalBatchBytes = 0,

        # Bytes already copied before this file
        [long]$BytesBefore = 0
    )

    $fileName  = Split-Path $Source -Leaf
    $fileBytes = (Get-Item $Source).Length
    $destPath  = if (Test-Path $Destination -PathType Container) {
                     Join-Path $Destination $fileName
                 } else { $Destination }

    $bufferSize = 4194304  # 4 MB
    $buffer     = New-Object byte[] $bufferSize

    $src = [System.IO.File]::OpenRead($Source)
    $dst = [System.IO.File]::Open($destPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)

    try {
        $copied = 0L
        while ($true) {
            $read = $src.Read($buffer, 0, $bufferSize)
            if ($read -eq 0) { break }
            $dst.Write($buffer, 0, $read)
            $copied += $read

            $fileMB    = [math]::Round($copied   / 1MB, 0)
            $fileTotMB = [math]::Round($fileBytes / 1MB, 0)
            $pct = if ($TotalBatchBytes -gt 0) {
                       [math]::Round(($BytesBefore + $copied) / $TotalBatchBytes * 100, 0)
                   } else {
                       [math]::Round($copied / $fileBytes * 100, 0)
                   }

            Write-Progress -Activity "Copying WIM Files" `
                           -Status "${fileName}  ${fileMB} / ${fileTotMB} MB" `
                           -PercentComplete $pct
        }
    }
    finally {
        $src.Close()
        $dst.Close()
    }
}

function Copy-ModelWIMs {
    <#
    .SYNOPSIS
        Copies model-specific WIM files (and optional drivers) to the USB drive.

        Source directory structure under WimSourceDirectory:
          <AnyName>\                  - directory name is used as model string (no model.ini)
          <AnyName>\model.ini         - overrides model string (ModelString=<value>)
          <AnyName>\Drivers\          - drivers to inject; copied to I:\Drivers\<ModelString>\

        Display format in the selection list:
          B9450FA (Pure)                       - no model.ini, no Drivers : dir name = model, factory capture
          B9450FA (Pure + Drivers)             - no model.ini, Drivers present : dir name = model, vanilla OS
          EXPERTBOOK_B9450FA_OEM [B9450FA]     - model.ini present, no Drivers : descriptive name, factory capture
          EXPERTBOOK_B9450FA_PRO [B9450FA] (Drivers) - model.ini present, Drivers present : descriptive name, vanilla OS
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DriveLetter,

        [Parameter(Mandatory)]
        [string]$WimSourceDirectory
    )

    Write-Section "Copy Model WIMs"

    # Validate WIM source directory
    if (-not (Test-Path $WimSourceDirectory -PathType Container)) {
        Write-Log "WIM source directory not found: $WimSourceDirectory" -Level Error
        return $false
    }

    # Enumerate all subdirectories - no name-length restriction (dirs can be descriptive)
    $modelDirs = Get-ChildItem -Path $WimSourceDirectory -Directory

    if ($modelDirs.Count -eq 0) {
        Write-Log "No directories found in: $WimSourceDirectory" -Level Warning
        return $false
    }

    # Build enriched list: resolve model name eagerly so it appears in the label
    $entries = foreach ($dir in $modelDirs) {
        $hasModelIni = Test-Path (Join-Path $dir.FullName "model.ini")
        $hasDrivers  = Test-Path (Join-Path $dir.FullName "Drivers") -PathType Container

        # A directory without model.ini must have a name that IS a valid model string
        # (typically 6-10 characters, no separators). Warn and skip if it looks descriptive.
        if (-not $hasModelIni -and $dir.Name.Length -gt 10) {
            Write-Log "Skipping '$($dir.Name)': name is too long for a model string and no model.ini found." -Level Warning
            continue
        }

        # Resolve ModelString now so we can embed it in the label
        $modelString = $dir.Name   # default: directory name IS the model
        if ($hasModelIni) {
            $modelIniPath = Join-Path $dir.FullName "model.ini"
            $modelLine = Get-Content $modelIniPath -ErrorAction SilentlyContinue |
                         Where-Object { $_ -match '^ModelString=' } |
                         Select-Object -First 1
            if ($modelLine) {
                $parsed = ($modelLine -split '=', 2)[1].Trim()
                if ($parsed) { $modelString = $parsed }
            }
        }

        # Build display label
        #   No model.ini  -> (Pure) or (Pure + Drivers)   : dir name IS the model string
        #   Has model.ini -> [RealModel] with optional (Drivers) : descriptive dir name
        $label = if (-not $hasModelIni) {
            $tag = if ($hasDrivers) { " (Pure + Drivers)" } else { " (Pure)" }
            "$($dir.Name)$tag"
        } else {
            $tag = if ($hasDrivers) { " (Drivers)" } else { "" }
            "$($dir.Name) [$modelString]$tag"
        }

        [PSCustomObject]@{
            Label       = $label
            DirName     = $dir.Name
            Path        = $dir.FullName
            HasModelIni = $hasModelIni
            HasDrivers  = $hasDrivers
            ModelString = $modelString
        }
    }

    Write-Host ""
    Write-Host "  Available Models:" -ForegroundColor Cyan
    $selectedLabel = Select-FromList -Items ($entries | Select-Object -ExpandProperty Label) -Prompt "Select model"
    $selected = $entries | Where-Object { $_.Label -eq $selectedLabel }

    $modelString = $selected.ModelString
    Write-Log "Selected: $($selected.DirName)  ->  ModelString: $modelString" -Level Info

    $destination = "${DriveLetter}\"

    # Update config.ini with resolved model name and auto-detected edition index
    $configFile = Join-Path $destination "config.ini"
    if (Test-Path $configFile) {
        $content = Get-Content $configFile

        # Always update ModelString
        $content = $content -replace '^ModelString=.*', "ModelString=$modelString"
        Write-Log "Updating config.ini: ModelString=$modelString" -Level Info

        # Auto-detect Pro edition index from the OS WIM (source, before copy)
        $osWimSource = Get-ChildItem -Path $selected.Path -File |
                       Where-Object { $_.Name -match '^.*-OS\.wim$' -or $_.Name -match '^.*-OS\.swm$' } |
                       Select-Object -First 1

        if ($osWimSource) {
            try {
                $images = Get-WindowsImage -ImagePath $osWimSource.FullName -ErrorAction Stop
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
                    Write-Log "No Pro edition found in WIM ($($images.Count) editions present) — EditionIndex left as-is." -Level Warning
                    Write-Log "Available editions: $(($images | Select-Object -ExpandProperty ImageName) -join ', ')" -Level Warning
                }
            }
            catch {
                Write-Log "Could not read WIM image info: $_" -Level Warning
                Write-Log "EditionIndex left as-is in config.ini." -Level Warning
            }
        } else {
            Write-Log "No OS WIM found in source — EditionIndex left as-is." -Level Warning
        }

        Set-Content -Path $configFile -Value $content
    }

    # Copy WIM/SWM files from the directory root only (exclude Drivers\ subfolder)
    $wimFiles = Get-ChildItem -Path $selected.Path -File |
                Where-Object { $_.Extension -in @('.wim', '.swm') }

    if ($wimFiles.Count -eq 0) {
        Write-Log "No WIM/SWM files found in $($selected.Path)" -Level Warning
        return $false
    }

    $totalSizeBytes = ($wimFiles | Measure-Object -Property Length -Sum).Sum
    $totalSizeMB    = [math]::Round($totalSizeBytes / 1MB, 0)
    Write-Log "Copying $($wimFiles.Count) WIM/SWM file(s) (${totalSizeMB} MB total)..." -Level Info

    $copiedBytes = 0L
    foreach ($file in $wimFiles) {
        Copy-FileWithProgress -Source $file.FullName -Destination $destination `
                              -TotalBatchBytes $totalSizeBytes -BytesBefore $copiedBytes
        $copiedBytes += $file.Length
    }

    Write-Progress -Activity "Copying WIM Files" -Completed

    $totalSizeGB = [math]::Round($totalSizeBytes / 1GB, 2)
    Write-Log "Copied $($wimFiles.Count) WIM/SWM file(s) (${totalSizeGB} GB total)." -Level Success

    # Copy drivers if present - required when Drivers\ folder exists
    if ($selected.HasDrivers) {
        $driverSource = Join-Path $selected.Path "Drivers"
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

$config = Import-PowerShellDataFile -Path $configPath
$PE_WIM = $config["WinPE_WIM_Location"]
$DEV_ROOT_WIM = $config["Device_Root_WIM_Location"]

Write-Log "WinPE WIM: $PE_WIM" -Level Info
Write-Log "Device WIM Root: $DEV_ROOT_WIM" -Level Info

# Validate WinPE WIM path (required for PE operations)
$peWimValid = Test-Path -Path $PE_WIM -PathType Leaf
if ($peWimValid) {
    Write-Log "WinPE WIM file found." -Level Success
} else {
    Write-Log "WinPE WIM file not found (required for PE installation)." -Level Warning
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

if ($imagesDrive -and (Read-YesNo "Do you want to copy model WIM files to ${imagesDrive}?")) {

    # Validate Device WIM root only when needed
    if (-not (Test-Path -Path $DEV_ROOT_WIM -PathType Container)) {
        Write-Log "Device WIM root directory not found: $DEV_ROOT_WIM" -Level Error
        Write-Log "Please update Device_Root_WIM_Location in config.psd1" -Level Warning
    } else {
        Copy-ModelWIMs -DriveLetter $imagesDrive -WimSourceDirectory $DEV_ROOT_WIM
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
                            $rst = Find-RSTDriverPath -DriverDir $srcDriverDir
                            if ($rst) {
                                $rstEntries += [PSCustomObject]@{
                                    Label = "$($dir.Name)  [from repository]"
                                    Path  = $rst
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
                Write-Log "$(if ($rstEntries.Count -eq 1) { 'Driver source found' } else { 'Multiple driver sources found' }) — select one to inject:" -Level Info
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
