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

    # Create mount directory
    if (-not (Test-Path $mountDir)) {
        New-Item -Path $mountDir -ItemType Directory -Force | Out-Null
    }

    try {
        # Mount boot.wim
        Write-Log "Mounting boot.wim..." -Level Info
        Write-Host ""
        & dism.exe /Mount-Image /ImageFile:"$bootWim" /Index:1 /MountDir:"$mountDir"

        if ($LASTEXITCODE -ne 0) {
            Write-Log "Failed to mount boot.wim (exit code: $LASTEXITCODE)" -Level Error
            return $false
        }

        Write-Log "boot.wim mounted successfully." -Level Success

        # Inject drivers
        Write-Log "Injecting drivers from: $DriverPath" -Level Info
        Write-Host ""
        $dismCmd = "dism.exe /Image:`"$mountDir`" /Add-Driver /Driver:`"$DriverPath`" /Recurse"
        Write-Log "Running: $dismCmd" -Level Info
        cmd /c $dismCmd

        if ($LASTEXITCODE -ne 0) {
            Write-Log "Warning: Some drivers may have failed to inject (exit code: $LASTEXITCODE)" -Level Warning
        } else {
            Write-Log "Drivers injected successfully." -Level Success
        }

        # Unmount and commit
        Write-Log "Unmounting and committing changes..." -Level Info
        Write-Host ""
        cmd /c "dism.exe /Unmount-Image /MountDir:`"$mountDir`" /Commit"

        if ($LASTEXITCODE -ne 0) {
            Write-Log "Failed to unmount boot.wim (exit code: $LASTEXITCODE)" -Level Error
            # Try to discard changes
            cmd /c "dism.exe /Unmount-Image /MountDir:`"$mountDir`" /Discard"
            return $false
        }

        Write-Log "WinPE drivers injected and saved successfully." -Level Success
        return $true
    }
    catch {
        Write-Log "Error injecting drivers: $_" -Level Error
        # Try to cleanup mount
        cmd /c "dism.exe /Unmount-Image /MountDir:`"$mountDir`" /Discard" 2>$null
        return $false
    }
    finally {
        # Cleanup mount directory if empty
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
        Copies model-specific WIM files to the USB drive
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

    # Find model directories (typically 6-8 characters, but accept any reasonable length)
    $modelDirs = Get-ChildItem -Path $WimSourceDirectory -Directory | Where-Object { $_.Name.Length -ge 5 -and $_.Name.Length -le 10 }

    if ($modelDirs.Count -eq 0) {
        Write-Log "No model directories found (expecting 5-10 character folder names)." -Level Warning
        return $false
    }

    Write-Host ""
    Write-Host "  Available Models:" -ForegroundColor Cyan
    $modelNames = $modelDirs | Select-Object -ExpandProperty Name

    $selectedModel = Select-FromList -Items $modelNames -Prompt "Select model"

    Write-Log "Selected model: $selectedModel" -Level Info

    $sourceDir = Join-Path $WimSourceDirectory $selectedModel
    $destination = "${DriveLetter}\"

    # Update config.ini with model name
    $configFile = Join-Path $destination "config.ini"
    if (Test-Path $configFile) {
        Write-Log "Updating config.ini with ModelString=$selectedModel" -Level Info
        $content = Get-Content $configFile
        $content = $content -replace '^ModelString=.*', "ModelString=$selectedModel"
        Set-Content -Path $configFile -Value $content
    }

    # Copy WIM and SWM files
    $wimFiles = Get-ChildItem -Path $sourceDir -Include "*.wim", "*.swm" -Recurse

    if ($wimFiles.Count -eq 0) {
        Write-Log "No WIM/SWM files found in $sourceDir" -Level Warning
        return $false
    }

    Write-Log "Copying $($wimFiles.Count) WIM/SWM files..." -Level Info

    $counter = 0
    $totalSize = ($wimFiles | Measure-Object -Property Length -Sum).Sum
    $copiedSize = 0

    foreach ($file in $wimFiles) {
        $counter++
        $percentComplete = [math]::Round(($copiedSize / $totalSize) * 100, 0)
        $sizeMB = [math]::Round($file.Length / 1MB, 0)

        Write-Progress -Activity "Copying WIM Files" -Status "$($file.Name) (${sizeMB}MB)" -PercentComplete $percentComplete

        Copy-Item -Path $file.FullName -Destination $destination -Force
        $copiedSize += $file.Length
    }

    Write-Progress -Activity "Copying WIM Files" -Completed

    $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
    Write-Log "Copied $counter files (${totalSizeGB}GB total)." -Level Success

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
# Step 3: Inject Drivers into WinPE (Optional)
# ============================================================
Write-Section "Step 3: Inject Drivers into WinPE (Optional)"

Write-Host ""
Write-Log "You can inject storage drivers (e.g., Intel RST) to enable disk detection in WinPE." -Level Info
Write-Host ""

if (Read-YesNo "Do you want to inject drivers into WinPE boot.wim?") {
    # If WinPE drive not already set, ask user to select it
    if (-not $winPEDrive) {
        Write-Host ""
        Write-Log "Select the WinPE partition (FAT32 partition with sources\boot.wim):" -Level Info
        $selectedPE = Select-DriveLetter -Purpose "WinPE partition"
        if ($selectedPE) {
            $winPEDrive = $selectedPE.DriveLetter
        }
    }

    if ($winPEDrive) {
        # Verify boot.wim exists
        $bootWimPath = Join-Path $winPEDrive "sources\boot.wim"
        if (-not (Test-Path $bootWimPath)) {
            Write-Log "boot.wim not found at $bootWimPath - is this the correct WinPE partition?" -Level Error
        } else {
            # Ask for Images drive letter to find drivers
            Write-Host ""
            $imagesLetter = Read-KeyPress "Enter the Images drive letter (e.g., I): "

            if ($imagesLetter) {
                $imagesDriveLetter = "${imagesLetter}:"
                $driversRoot = Join-Path $imagesDriveLetter "Drivers"

                if (-not (Test-Path $driversRoot -PathType Container)) {
                    Write-Log "Drivers folder not found at $driversRoot" -Level Error
                } else {
                    # Get first model directory under Drivers
                    $modelDir = Get-ChildItem -Path $driversRoot -Directory | Select-Object -First 1

                    if (-not $modelDir) {
                        Write-Log "No model directory found under $driversRoot" -Level Error
                    } else {
                        # Look for Rapid Storage directory
                        $rstPath = Join-Path $modelDir.FullName "Rapid Storage"

                        if (-not (Test-Path $rstPath -PathType Container)) {
                            Write-Log "Rapid Storage folder not found at $rstPath" -Level Error
                            Write-Log "Looking for alternative storage driver folders..." -Level Info

                            # Try to find any folder with "Storage" or "RST" in name
                            $altDriver = Get-ChildItem -Path $modelDir.FullName -Directory -Recurse |
                                Where-Object { $_.Name -match 'Storage|RST|IRST' } |
                                Select-Object -First 1

                            if ($altDriver) {
                                $rstPath = $altDriver.FullName
                                Write-Log "Found alternative: $rstPath" -Level Info
                            } else {
                                Write-Log "No storage driver folder found." -Level Error
                                $rstPath = $null
                            }
                        }

                        if ($rstPath) {
                            Write-Log "Using driver path: $rstPath" -Level Info
                            Write-Log "Model: $($modelDir.Name)" -Level Info
                            Add-WinPEDrivers -WinPEDriveLetter $winPEDrive -DriverPath $rstPath
                        }
                    }
                }
            } else {
                Write-Log "No drive letter provided. Skipping." -Level Warning
            }
        }
    } else {
        Write-Log "No WinPE partition selected. Skipping." -Level Warning
    }
}

# ============================================================
# Step 4: Copy Script Files (Optional)
# ============================================================
Write-Section "Step 4: Copy Script Files (Optional)"

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
# Step 5: Copy Model WIMs (Optional)
# ============================================================
Write-Section "Step 5: Copy Model WIMs (Optional)"

if ($imagesDrive -and (Read-YesNo "Do you want to copy model WIM files to $imagesDrive ?")) {

    # Validate Device WIM root only when needed
    if (-not (Test-Path -Path $DEV_ROOT_WIM -PathType Container)) {
        Write-Log "Device WIM root directory not found: $DEV_ROOT_WIM" -Level Error
        Write-Log "Please update Device_Root_WIM_Location in config.psd1" -Level Warning
    } else {
        Copy-ModelWIMs -DriveLetter $imagesDrive -WimSourceDirectory $DEV_ROOT_WIM
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
