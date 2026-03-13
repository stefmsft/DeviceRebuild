<#
.SYNOPSIS
    DeviceRebuild shared module — utility functions used by ProduceKey.ps1,
    ExtractWim.ps1, and future tooling scripts.
#>

# ============================================================
# Module-scope logging state — set by the calling script via Initialize-Logging
# ============================================================
$Script:LogEnabled = $false
$Script:LogFile    = $null

function Initialize-Logging {
    <#
    .SYNOPSIS
        Configures the module logging state for the calling script session.
        Must be called once at startup before any Write-Log output.
    #>
    param(
        [bool]$Enabled,
        [string]$LogFile
    )
    $Script:LogEnabled = $Enabled
    $Script:LogFile    = $LogFile
}

# ============================================================
# Logging / UI
# ============================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"

    if ($Script:LogEnabled) {
        Add-Content -Path $Script:LogFile -Value $logMessage
    }

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
# User Interaction
# ============================================================
function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $false
    )

    $defaultHint = if ($Default) { "[Y/n]" } else { "[y/N]" }
    Write-Host "$Prompt $defaultHint : " -NoNewline -ForegroundColor Yellow
    $response = Read-Host

    if ([string]::IsNullOrWhiteSpace($response)) { return $Default }
    return $response -match '^[Yy]'
}

function Read-KeyPress {
    param([string]$Prompt)

    Write-Host $Prompt -NoNewline -ForegroundColor Yellow
    $key = [System.Console]::ReadKey($true)
    Write-Host $key.KeyChar

    if ($key.Key -eq [ConsoleKey]::Escape) { return $null }
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
# USB Drive Helpers
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
# DISM helpers
# ============================================================
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
            } else {
                Write-Log "Failed to discard stale mount (exit code: $($p.ExitCode)) — forcing Cleanup-Wim." -Level Warning
            }
            $cleaned = $true
        }

        if ($cleaned) {
            Write-Log "Running Dism /Cleanup-Wim to clear residual mount metadata..." -Level Info
            & dism.exe /Cleanup-Wim 2>&1 | Out-Null
            Write-Log "DISM cleanup complete." -Level Info
        }

        return $cleaned
    }
    catch {
        Write-Log "Error querying mounted images: $_" -Level Error
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

    if ($alt) { return $alt.FullName } else { return $null }
}

function Get-WimDestName {
    <#
    .SYNOPSIS
        Computes the destination filename for a WIM/SWM on the USB key.
        Convention: <TargetModel>-<PartitionType>.ext
        - "GENERIC-RECOVERY.wim"  -> "B1403CVA-RECOVERY.wim"
        - "RECOVERY.wim"          -> "B1403CVA-RECOVERY.wim"
        - "GENERIC-OS2.swm"       -> "B1403CVA-OS2.swm"
    #>
    param(
        [Parameter(Mandatory)] [string]$SourceName,
        [Parameter(Mandatory)] [string]$TargetModel
    )

    $ext  = [System.IO.Path]::GetExtension($SourceName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($SourceName)

    # Extract partition type: everything after the last '-', or the whole name if no '-'
    $partType = if ($base -match '-([^-]+)$') { $Matches[1] } else { $base }

    return "$TargetModel-$partType$ext"
}

function Resolve-ModelDirectory {
    <#
    .SYNOPSIS
        Resolves a model name (from a LNK filename) to the actual directory path
        under WimSourceDirectory. Supports both exact directory name matches and
        ModelString prefix matches (e.g. "GENERIC" resolves to "GENERIC - W11-25H2 - US").
    .OUTPUTS
        Full path string, or $null if not found.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ModelName,

        [Parameter(Mandatory)]
        [string]$WimSourceDirectory
    )

    # Exact match first
    $exact = Join-Path $WimSourceDirectory $ModelName
    if (Test-Path $exact -PathType Container) { return $exact }

    # Prefix match: directory whose name starts with "$ModelName - " or equals "$ModelName"
    $match = Get-ChildItem -Path $WimSourceDirectory -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -eq $ModelName -or $_.Name -match "^$([regex]::Escape($ModelName)) - " } |
             Select-Object -First 1

    if ($match) { return $match.FullName }
    return $null
}

function Resolve-DriverLnk {
    <#
    .SYNOPSIS
        Resolves the effective driver directory for a model, following a
        LNK-<model>.txt redirect file if one exists inside the Drivers\ folder.
        LNK targets resolve against Shared_Drivers_Location (<model> subfolder),
        keeping shared drivers independent from model directories.
    .OUTPUTS
        PSCustomObject { Path; LinkedModel } where LinkedModel is $null when
        no redirect is in effect, or $null if the Drivers folder is absent /
        the LNK target does not exist.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DriversDir,           # The model's Drivers\ folder path

        [Parameter(Mandatory)]
        [string]$SharedDriversLocation # Shared_Drivers_Location from config.psd1
    )

    if (-not (Test-Path $DriversDir -PathType Container)) { return $null }

    $lnk = Get-ChildItem -Path $DriversDir -File -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -match '^LNK-(.+)\.txt$' } |
           Select-Object -First 1

    if ($lnk) {
        $linkedModel      = [regex]::Match($lnk.Name, '^LNK-(.+)\.txt$').Groups[1].Value
        $linkedDriversDir = Join-Path $SharedDriversLocation $linkedModel
        if (Test-Path $linkedDriversDir -PathType Container) {
            return [PSCustomObject]@{ Path = $linkedDriversDir; LinkedModel = $linkedModel }
        }
        Write-Log "Driver LNK '$($lnk.Name)' target not found: $linkedDriversDir" -Level Warning
        Write-Log "  Expected: $SharedDriversLocation\$linkedModel\" -Level Warning
        return $null
    }

    return [PSCustomObject]@{ Path = $DriversDir; LinkedModel = $null }
}

# ============================================================
# File copy
# ============================================================
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
        [long]$BytesBefore = 0,

        # Activity label shown in the progress bar
        [string]$Activity = "Copying"
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

            Write-Progress -Activity $Activity `
                           -Status "${fileName}  ${fileMB} / ${fileTotMB} MB" `
                           -PercentComplete $pct
        }
    }
    finally {
        $src.Close()
        $dst.Close()
    }
}
