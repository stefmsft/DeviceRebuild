#Requires -RunAsAdministrator
<#
.SYNOPSIS
    ExtractWim.ps1 - Extracts Pro and Pro Education editions from a Windows ISO.

.DESCRIPTION
    Mounts a Windows ISO, identifies the Windows version from the build number,
    and exports the Pro and Pro Education editions into a version-named WIM file
    under Windows_WIM_Root (configured in config.psd1).

    Output: <Windows_WIM_Root>\<Version>\<Version>.wim
    Example: C:\Scratch\WindowsWIMs\W11-24H2\W11-24H2.wim

.PARAMETER IsoPath
    Full path to the Windows ISO file. If omitted, a file browser dialog opens.

.PARAMETER Log
    Enable logging to file.

.NOTES
    Requires: Administrator privileges, Windows ADK (DISM)
#>
param(
    [string]$IsoPath,
    [switch]$Log
)

# ============================================================
# Module + logging init
# ============================================================
Import-Module (Join-Path $PSScriptRoot "DeviceRebuild.psm1") -Force

$Script:LogEnabled = $Log.IsPresent
$Script:LogFile    = if ($Script:LogEnabled) { "ExtractWim_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" } else { $null }
Initialize-Logging -Enabled $Script:LogEnabled -LogFile $Script:LogFile

# ============================================================
# Build-number → version-name lookup table
# ============================================================
$BuildToVersion = @{
    10240 = 'W10-1507'
    10586 = 'W10-1511'
    14393 = 'W10-1607'
    15063 = 'W10-1703'
    16299 = 'W10-1709'
    17134 = 'W10-1803'
    17763 = 'W10-1809'
    18362 = 'W10-1903'
    18363 = 'W10-1909'
    19041 = 'W10-2004'
    19042 = 'W10-20H2'
    19043 = 'W10-21H1'
    19044 = 'W10-21H2'
    19045 = 'W10-22H2'
    22000 = 'W11-21H2'
    22621 = 'W11-22H2'
    22631 = 'W11-23H2'
    26100 = 'W11-24H2'
    26200 = 'W11-25H2'
}

# ============================================================
# Load configuration
# ============================================================
Write-Banner "ExtractWim - Windows ISO Edition Extractor"

$configPath = Join-Path $PSScriptRoot "config.psd1"
if (-not (Test-Path $configPath)) {
    Write-Log "config.psd1 not found at: $configPath" -Level Error
    exit 1
}

$config        = Import-PowerShellDataFile -Path $configPath
$WIN_WIM_ROOT  = $config["Windows_WIM_Root"]

if ([string]::IsNullOrWhiteSpace($WIN_WIM_ROOT)) {
    Write-Log "Windows_WIM_Root is not set in config.psd1" -Level Error
    exit 1
}

Write-Log "Windows WIM root: $WIN_WIM_ROOT" -Level Info

# ============================================================
# Resolve ISO path
# ============================================================
Write-Section "Select ISO"

if ([string]::IsNullOrWhiteSpace($IsoPath)) {
    # Try GUI file picker first; fall back to typed input if unavailable
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog        = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title  = "Select Windows ISO"
        $dialog.Filter = "ISO files (*.iso)|*.iso|All files (*.*)|*.*"
        if ($dialog.ShowDialog() -eq 'OK') {
            $IsoPath = $dialog.FileName
        }
    }
    catch {
        # Running headless — fall back to typed input
    }
}

if ([string]::IsNullOrWhiteSpace($IsoPath)) {
    Write-Host "Enter path to Windows ISO: " -NoNewline -ForegroundColor Yellow
    $IsoPath = Read-Host
}

$IsoPath = $IsoPath.Trim('"')

if (-not (Test-Path $IsoPath -PathType Leaf)) {
    Write-Log "ISO not found: $IsoPath" -Level Error
    exit 1
}

Write-Log "ISO: $IsoPath" -Level Info

# ============================================================
# Mount ISO
# ============================================================
Write-Section "Mount ISO"
Write-Log "Mounting ISO..." -Level Info

$mountResult  = Mount-DiskImage -ImagePath $IsoPath -PassThru -ErrorAction Stop
$isoDrive     = ($mountResult | Get-Volume).DriveLetter + ":"
Write-Log "ISO mounted at $isoDrive" -Level Success

try {

    # ============================================================
    # Locate install.wim / install.esd
    # ============================================================
    $installWim = Join-Path $isoDrive "sources\install.wim"
    $installEsd = Join-Path $isoDrive "sources\install.esd"

    if (Test-Path $installWim) {
        $sourceWim = $installWim
        Write-Log "Found: sources\install.wim" -Level Info
    } elseif (Test-Path $installEsd) {
        $sourceWim = $installEsd
        Write-Log "Found: sources\install.esd" -Level Info
        Write-Log "Note: .esd files are compressed — export will be slower." -Level Warning
    } else {
        Write-Log "Neither install.wim nor install.esd found under $isoDrive\sources\" -Level Error
        exit 1
    }

    # ============================================================
    # Read image list
    # ============================================================
    Write-Section "Read Image Info"
    Write-Log "Reading image list from WIM..." -Level Info

    $allImages = Get-WindowsImage -ImagePath $sourceWim -ErrorAction Stop
    Write-Log "Found $($allImages.Count) edition(s) in WIM:" -Level Info
    foreach ($img in $allImages) {
        Write-Log "  Index $($img.ImageIndex): $($img.ImageName)" -Level Info
    }

    # ============================================================
    # Detect Windows version from build number
    # ============================================================
    Write-Section "Detect Windows Version"

    $firstImage  = Get-WindowsImage -ImagePath $sourceWim -Index 1 -ErrorAction Stop
    $versionObj  = [System.Version]$firstImage.Version
    $buildNumber = $versionObj.Build
    Write-Log "Build number: $buildNumber" -Level Info

    if ($BuildToVersion.ContainsKey($buildNumber)) {
        $versionName = $BuildToVersion[$buildNumber]
        Write-Log "Detected version: $versionName" -Level Success
    } else {
        Write-Log "Build $buildNumber is not in the known version table." -Level Warning
        Write-Host "Enter version name for this build (e.g. W11-25H2): " -NoNewline -ForegroundColor Yellow
        $versionName = Read-Host
        $versionName = $versionName.Trim()
        if ([string]::IsNullOrWhiteSpace($versionName)) {
            Write-Log "No version name provided. Aborting." -Level Error
            exit 1
        }
    }

    # ============================================================
    # Filter Pro and Pro Education editions
    # ============================================================
    $toExport = $allImages | Where-Object {
        $_.ImageName -match '\bPro$' -or $_.ImageName -match '\bPro Education$'
    }

    if ($toExport.Count -eq 0) {
        Write-Log "No Pro or Pro Education editions found in this WIM." -Level Error
        Write-Log "Available editions: $(($allImages | Select-Object -ExpandProperty ImageName) -join ', ')" -Level Warning
        exit 1
    }

    Write-Log "Editions to export:" -Level Info
    foreach ($img in $toExport) {
        Write-Log "  [$($img.ImageIndex)] $($img.ImageName)" -Level Info
    }

    # ============================================================
    # Check destination / overwrite
    # ============================================================
    Write-Section "Export"

    $destDir = Join-Path $WIN_WIM_ROOT $versionName
    $destWim = Join-Path $destDir "$versionName.wim"

    if (Test-Path $destWim) {
        Write-Log "$destWim already exists." -Level Warning
        if (-not (Read-YesNo "Overwrite existing $versionName.wim?")) {
            Write-Log "Export cancelled by user." -Level Warning
            exit 0
        }
        Remove-Item $destWim -Force
        Write-Log "Existing file removed." -Level Info
    }

    if (-not (Test-Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        Write-Log "Created directory: $destDir" -Level Info
    }

    # ============================================================
    # Export editions
    # ============================================================
    $exportedCount = 0

    foreach ($img in $toExport) {
        Write-Log "Exporting '$($img.ImageName)' (index $($img.ImageIndex))..." -Level Info
        Write-Log "(This may take several minutes)" -Level Warning
        Write-Host ""

        # dism /Export-Image shows its own progress bar and appends to an existing WIM
        & dism.exe /Export-Image `
            "/SourceImageFile:$sourceWim" `
            "/SourceIndex:$($img.ImageIndex)" `
            "/DestinationImageFile:$destWim" `
            /Compress:max `
            /CheckIntegrity

        if ($LASTEXITCODE -eq 0) {
            Write-Log "'$($img.ImageName)' exported successfully." -Level Success
            $exportedCount++
        } else {
            Write-Log "Export failed for '$($img.ImageName)' (exit code: $LASTEXITCODE)" -Level Error
        }

        Write-Host ""
    }

    # ============================================================
    # Summary
    # ============================================================
    Write-Banner "Complete"

    if ($exportedCount -gt 0) {
        $sizeMB = [math]::Round((Get-Item $destWim).Length / 1MB, 0)
        Write-Log "Exported $exportedCount edition(s) to: $destWim  (${sizeMB} MB)" -Level Success
    } else {
        Write-Log "No editions were exported successfully." -Level Error
    }

    if ($Script:LogEnabled) {
        Write-Log "Log saved to: $Script:LogFile" -Level Info
    }
}
finally {
    # Always unmount the ISO
    Write-Log "Unmounting ISO..." -Level Info
    Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue | Out-Null
    Write-Log "ISO unmounted." -Level Info
}
