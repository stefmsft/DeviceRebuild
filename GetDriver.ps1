<#
.SYNOPSIS
    GetDriver.ps1 - Automated SCCM Driver Package downloader for ASUS devices.

.DESCRIPTION
    Searches for and downloads SCCM Driver Packages from the ASUS Support site.
    Supports interactive and non-interactive (silent) modes.

.PARAMETER Model
    The device model name (e.g., BM1403 or BM1403CDA).

.PARAMETER OSVersion
    The target OS version (e.g., 24H2, 23H2). Defaults to 25H2.

.PARAMETER Version
    Specific version string (vxx.yy) to download.

.PARAMETER Destination
    Directory to save the download. Defaults to the current directory.

.PARAMETER AutoExpand
    Controls extraction. If provided ($true or $false), the script runs silently for this step.
    If $true, extracts and deletes the source ZIP.

.PARAMETER GetLatestVersion
    Returns only a JSON object with {Version, OSVersion} and exits.

.EXAMPLE
    .\GetDriver.ps1 -Model "BM1403" -GetLatestVersion
#>
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$Model,

    [string]$OSVersion = "25H2",

    [string]$Version,

    [string]$Destination = ".",

    [Parameter(Mandatory=$false)]
    [object]$AutoExpand,

    [switch]$GetLatestVersion
)

# ============================================================
# Core Functions (Callable from other scripts)
# ============================================================

function Get-AsusDriverPackages {
    param(
        [string]$ModelName,
        [string]$Pdid = "",
        [string]$OsId = "52"
    )

    $websites = @("global", "fr")
    $allDrivers = @()

    foreach ($site in $websites) {
        $apiUrl = "https://www.asus.com/support/api/product.asmx/GetPDDrivers?website=$site&model=$ModelName&pdid=$Pdid&osid=$OsId"
        
        try {
            $response = Invoke-RestMethod -Uri $apiUrl -Method Get -ErrorAction Stop
            if ($response.Result.Obj -and $response.Result.Obj.Count -gt 0) {
                foreach ($category in $response.Result.Obj) {
                    if ($null -eq $category.Files) { continue }
                    foreach ($item in $category.Files) {
                        $allDrivers += [PSCustomObject]@{
                            Title       = $item.Title
                            Version     = $item.Version
                            ReleaseDate = $item.ReleaseDate
                            DownloadUrl = $item.DownloadUrl.Global
                            Size        = $item.FileSize
                            Description = $item.Description
                        }
                    }
                }
                if ($allDrivers.Count -gt 0) { break }
            }
        } catch { }
    }
    return $allDrivers
}

function Get-AsusProductInfo {
    param([string]$Keyword)

    $testDrivers = Get-AsusDriverPackages -ModelName $Keyword
    if ($testDrivers.Count -gt 0) {
        return [PSCustomObject]@{ Pdid = ""; ModelName = $Keyword.ToUpper() }
    }

    $searchUrl = "https://www.asus.com/support/api/product.asmx/SearchProduct?keyword=$Keyword&website=global&type=1"
    try {
        $response = Invoke-RestMethod -Uri $searchUrl -Method Get -ErrorAction Stop
        if ($response.Result.Obj -and $response.Result.Obj.Count -gt 0) {
            $product = $response.Result.Obj[0]
            return [PSCustomObject]@{ Pdid = $product.PDID; ModelName = $product.PDName }
        }
    } catch { }
    return $null
}

function Expand-DriverPackage {
    param(
        [Parameter(Mandatory)] [string]$ZipPath,
        [string]$DestDir
    )

    if (-not (Test-Path $ZipPath)) { Write-Error "ZIP file not found: $ZipPath"; return $false }

    if (-not $DestDir) {
        $DestDir = [System.IO.Path]::Combine((Split-Path $ZipPath), [System.IO.Path]::GetFileNameWithoutExtension($ZipPath))
    }

    Write-Host "Extracting to: $DestDir..." -ForegroundColor Cyan
    if (-not (Test-Path $DestDir)) { New-Item -ItemType Directory -Path $DestDir -Force | Out-Null }

    try {
        Expand-Archive -Path $ZipPath -DestinationPath $DestDir -Force
        Write-Host "Extraction complete. Deleting ZIP file..." -ForegroundColor Green
        Remove-Item $ZipPath -Force
        return $true
    } catch {
        Write-Error "Failed to extract archive: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================
# Main Execution Logic
# ============================================================

# 1. Resolve Model (Quietly if JSON requested)
if (-not $GetLatestVersion) { Write-Host "Searching for model: $Model..." -NoNewline }
$product = Get-AsusProductInfo -Keyword $Model
if (-not $product) {
    if (-not $GetLatestVersion) { Write-Host " Not found." -ForegroundColor Red }
    exit 1
}
if (-not $GetLatestVersion) { Write-Host " Found: $($product.ModelName)" -ForegroundColor Green }

# 2. Fetch Driver List
$drivers = Get-AsusDriverPackages -ModelName $product.ModelName -Pdid $product.Pdid

# 3. Filter for SCCM Packages
$sccmPacks = $drivers | Where-Object { $_.Title -match "SCCM" }
if ($sccmPacks.Count -eq 0) {
    if (-not $GetLatestVersion) { Write-Host "No SCCM packages found." -ForegroundColor Yellow }
    exit 0
}

# 4. Handle -GetLatestVersion (Output JSON only)
if ($GetLatestVersion) {
    $targetPacks = $sccmPacks
    if ($PSBoundParameters.ContainsKey('OSVersion')) {
        $targetPacks = $sccmPacks | Where-Object { $_.Title -match $OSVersion }
    }
    
    if ($targetPacks.Count -gt 0) {
        $p = $targetPacks[0]
        $osMatch = [regex]::Match($p.Title, '\(([^)]+)\)')
        $osStr = if ($osMatch.Success) { $osMatch.Groups[1].Value } else { "Unknown" }
        @{ Version = $p.Version; OSVersion = $osStr } | ConvertTo-Json -Compress
    }
    exit 0
}

Write-Host "--- GetDriver - ASUS SCCM Pack Downloader ---" -ForegroundColor Magenta

# 5. Handle Selection
$selectedPack = $null
$IsInteractive = [Environment]::UserInteractive -and -not $PSBoundParameters.ContainsKey('AutoExpand') -and -not $Version -and -not $PSBoundParameters.ContainsKey('OSVersion')

if ($IsInteractive) {
    Write-Host "`nAvailable SCCM Driver Packages:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $sccmPacks.Count; $i++) {
        Write-Host "  [$($i + 1)] $($sccmPacks[$i].Title) ($($sccmPacks[$i].Version))"
    }
    do {
        $choice = Read-Host "`nSelect package (1-$($sccmPacks.Count))"
        $idx = $choice -as [int]
    } while ($idx -lt 1 -or $idx -gt $sccmPacks.Count)
    $selectedPack = $sccmPacks[$idx - 1]
    if ((Read-Host "Proceed with download? [Y/N]") -notmatch "^Y") { exit 0 }
} else {
    # Non-interactive: Filter by Version AND OSVersion if both provided
    $filtered = $sccmPacks
    if ($Version) {
        $filtered = $filtered | Where-Object { $_.Version -eq $Version }
    }
    if ($PSBoundParameters.ContainsKey('OSVersion')) {
        $filtered = $filtered | Where-Object { $_.Title -match $OSVersion }
    }
    
    $selectedPack = $filtered | Select-Object -First 1
    if (-not $selectedPack) { 
        Write-Error "No matching package found for Version='$Version' and OSVersion='$OSVersion'."
        exit 1
    }
}

# 6. Download
$fileName = (Split-Path $selectedPack.DownloadUrl -Leaf).Split('?')[0]
$outPath = Join-Path $Destination $fileName

Write-Host "`nDownloading: $fileName..." -ForegroundColor Cyan
if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
Invoke-WebRequest -Uri $selectedPack.DownloadUrl -OutFile $outPath

if (Test-Path $outPath) {
    Write-Host "Download complete." -ForegroundColor Green
    
    # 7. Extraction
    $doExpand = $false
    if ($PSBoundParameters.ContainsKey('AutoExpand')) {
        # Handles both -AutoExpand (switch style) and -AutoExpand $true/$false
        if ($AutoExpand -is [switch] -or $AutoExpand -is [bool]) { $doExpand = [bool]$AutoExpand }
        else { [bool]::TryParse($AutoExpand, [ref]$doExpand) | Out-Null }
    } else {
        $doExpand = (Read-Host "`nExtract drivers and delete ZIP? [Y/N]") -match "^Y"
    }

    if ($doExpand) { Expand-DriverPackage -ZipPath $outPath }
} else {
    Write-Error "Download failed."
}
