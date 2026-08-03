<#
.SYNOPSIS
    GetDriver.ps1 - Automated SCCM Driver Package downloader for ASUS devices.

.DESCRIPTION
    Searches for and downloads SCCM Driver Packages from the ASUS Support site.
    Supports interactive and non-interactive (silent) modes.

.PARAMETER Model
    The device model name (e.g., BM1403CDA).

.PARAMETER OSVersion
    The target OS version (e.g., 24H2, 23H2). Defaults to 25H2.

.PARAMETER Version
    Specific version string (vxx.yy) to download.

.PARAMETER Destination
    The root directory under which a directory named after the model will be created.
    If not provided, defaults to 'Shared_Drivers_Location' from config.psd1.

.PARAMETER AutoExpand
    Controls extraction. Defaults to $true. If $true, extracts and deletes the source ZIP.

.PARAMETER GetLatestVersion
    Returns only a JSON object with {Version, OSVersion} and exits.

.PARAMETER DryRun
    If specified, performs all logic (search, selection, etc.) but skips the actual download and extraction.
    Also disables interactive prompts by assuming defaults.

.EXAMPLE
    .\GetDriver.ps1 -Model "BM1403CDA" -GetLatestVersion
#>
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$Model,

    [string]$OSVersion = "25H2",

    [string]$Version,

    [string]$Destination,

    [Parameter(Mandatory=$false)]
    [object]$AutoExpand = $true,

    [switch]$GetLatestVersion,

    [switch]$DryRun
)

# ============================================================
# Load Configuration
# ============================================================
$ConfigPath = Join-Path $PSScriptRoot "config.psd1"
$config = if (Test-Path $ConfigPath) { 
    try {
        if (Get-Command Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
            Import-PowerShellDataFile -Path $ConfigPath
        } else {
            Invoke-Expression (Get-Content $ConfigPath -Raw)
        }
    } catch { $null }
} else { $null }

# ============================================================
# Core Functions
# ============================================================

function ConvertFrom-DriverApiResponse {
    param([object]$Response)
    $items = @()
    # Handle both old API (Result.Obj) and new API (Result.Result or direct Obj)
    $categories = if ($Response.Result.Obj) { $Response.Result.Obj }
                  elseif ($Response.Result.Result) { $Response.Result.Result }
                  else { $null }
    if (-not $categories) { return $items }
    foreach ($category in $categories) {
        $files = if ($category.Files) { $category.Files } elseif ($category.DriverList) { $category.DriverList } else { $null }
        if ($null -eq $files) { continue }
        foreach ($item in $files) {
            $dlUrl = if ($item.DownloadUrl.Global) { $item.DownloadUrl.Global }
                     elseif ($item.DownloadUrl) { $item.DownloadUrl }
                     else { "" }
            # Resolve relative URLs from the new API
            if ($dlUrl -and $dlUrl -notmatch "^https?://") {
                $dlUrl = "https://dlcdnets.asus.com$dlUrl"
            }
            $items += [PSCustomObject]@{
                Title       = $item.Title
                Version     = $item.Version
                ReleaseDate = $item.ReleaseDate
                DownloadUrl = $dlUrl
                Size        = $item.FileSize
                Description = $item.Description
            }
        }
    }
    return $items
}

function Get-AsusDriverPackages {
    param(
        [string]$ModelName,
        [string]$Pdid = "",
        [string]$OsId = "52"
    )

    $modelLower = $ModelName.ToLower()

    # --- New API (webapi/ProductV2) — used by ASUS website, works for "supportonly" models ---
    # pdid=99999 acts as a wildcard; cpu= filters by model
    foreach ($osParam in @($OsId, "")) {
        $url = "https://www.asus.com/support/webapi/ProductV2/GetPDDrivers?website=global&model=$modelLower&pdhashedid=&cpu=$ModelName&osid=$osParam&pdid=99999&siteID=www&sitelang="
        try {
            $r = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
            $items = ConvertFrom-DriverApiResponse -Response $r
            if ($items.Count -gt 0) { return $items }
        } catch { }
    }

    # --- Old API (product.asmx) — fallback for models resolved via product search ---
    $websites = @("global", "us", "fr", "uk", "tw")
    foreach ($osParam in @($OsId, "")) {
        foreach ($site in $websites) {
            $url = "https://www.asus.com/support/api/product.asmx/GetPDDrivers?website=$site&model=$ModelName&pdid=$Pdid&osid=$osParam"
            try {
                $r = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
                $items = ConvertFrom-DriverApiResponse -Response $r
                if ($items.Count -gt 0) { return $items }
            } catch { }
        }
    }

    return @()
}

function Get-AsusProductInfo {
    param([string]$Keyword)

    # 1. Direct Search
    $searchUrl = "https://www.asus.com/support/api/product.asmx/SearchProduct?keyword=$Keyword&website=global&type=1"
    try {
        $response = Invoke-RestMethod -Uri $searchUrl -Method Get -ErrorAction Stop
        if ($response.Result.Obj -and $response.Result.Obj.Count -gt 0) {
            # Try to find an exact match in the search results
            $match = $response.Result.Obj | Where-Object { $_.PDName -ieq $Keyword } | Select-Object -First 1
            if (-not $match) { $match = $response.Result.Obj[0] }
            
            return [PSCustomObject]@{ Pdid = $match.PDID; ModelName = $match.PDName }
        }
    } catch { }

    # 2. Try partial keyword if it's long (e.g. BM1403CDA -> BM1403)
    if ($Keyword.Length -gt 6) {
        $partial = $Keyword.Substring(0, 6)
        $searchUrl = "https://www.asus.com/support/api/product.asmx/SearchProduct?keyword=$partial&website=global&type=1"
        try {
            $response = Invoke-RestMethod -Uri $searchUrl -Method Get -ErrorAction Stop
            if ($response.Result.Obj -and $response.Result.Obj.Count -gt 0) {
                # Only return if we find an actual match — never silently fall back to Obj[0]
                $match = $response.Result.Obj | Where-Object { $_.PDName -match $Keyword } | Select-Object -First 1
                if ($match) {
                    return [PSCustomObject]@{ Pdid = $match.PDID; ModelName = $match.PDName }
                }
            }
        } catch { }
    }

    # 3. Fallback: try keyword as direct model name (no pdid)
    $testDrivers = Get-AsusDriverPackages -ModelName $Keyword
    if ($testDrivers.Count -gt 0) {
        return [PSCustomObject]@{ Pdid = ""; ModelName = $Keyword.ToUpper() }
    }

    # 4. Last resort: scrape PDID from the ASUS support page (handles "supportonly" models
    #    not indexed in the product search API but reachable via direct URL)
    $pageUrl = "https://www.asus.com/supportonly/$($Keyword.ToLower())/helpdesk_download/"
    try {
        $html = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -ErrorAction Stop
        $pdid = $null

        # ASUS embeds the PDID in various ways depending on page version
        foreach ($pattern in @('"pdid"\s*:\s*"?(\d+)"?', '"PDID"\s*:\s*"?(\d+)"?', 'pdid=(\d+)', "pdid%3D(\d+)")) {
            $m = [regex]::Match($html.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($m.Success) { $pdid = $m.Groups[1].Value; break }
        }

        if ($pdid) {
            $testDrivers = Get-AsusDriverPackages -ModelName $Keyword -Pdid $pdid
            if ($testDrivers.Count -gt 0) {
                return [PSCustomObject]@{ Pdid = $pdid; ModelName = $Keyword.ToUpper() }
            }
        }
    } catch { }

    return $null
}

function Expand-DriverPackage {
    param(
        [Parameter(Mandatory)] [string]$ZipPath,
        [string]$DestDir
    )

    if ($DryRun) {
        Write-Host "[DryRun] Would extract $ZipPath to $DestDir" -ForegroundColor Yellow
        return $true
    }

    if (-not (Test-Path $ZipPath)) { Write-Error "ZIP file not found: $ZipPath"; return $false }

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

# 1. Resolve Model
if (-not $GetLatestVersion) { Write-Host "Searching for model: $Model..." -NoNewline }
$product = Get-AsusProductInfo -Keyword $Model
if (-not $product) {
    if ($GetLatestVersion) { Write-Warning "Model '$Model' not found in ASUS product catalog." }
    else { Write-Host " Not found." -ForegroundColor Red }
    exit 1
}

# 2. Fetch Driver List
$drivers = Get-AsusDriverPackages -ModelName $product.ModelName -Pdid $product.Pdid

# 3. Filter for SCCM Packages
$sccmPacks = $drivers | Where-Object { $_.Title -match "SCCM|Pack" -and $_.Title -notmatch "MyASUS" }
if ($sccmPacks.Count -eq 0) {
    if ($GetLatestVersion) {
        Write-Warning "Model resolved as '$($product.ModelName)' but no SCCM packages found ($($drivers.Count) total drivers returned)."
    } else { Write-Host "No SCCM packages found." -ForegroundColor Yellow }
    exit 0
}

# 4. Handle -GetLatestVersion
if ($GetLatestVersion) {
    $targetPacks = $sccmPacks
    if ($PSBoundParameters.ContainsKey('OSVersion')) {
        $matches = $sccmPacks | Where-Object { $_.Title -match "\($OSVersion\)" }
        if (-not $matches) { $matches = $sccmPacks | Where-Object { $_.Title -match $OSVersion } }
        if ($matches) { $targetPacks = $matches }
    }

    if ($targetPacks.Count -gt 0) {
        $p = $targetPacks[0]
        $osMatch = [regex]::Match($p.Title, '\(([^)]+)\)')
        $osStr = if ($osMatch.Success) { $osMatch.Groups[1].Value } else { "Unknown" }
        @{ Version = $p.Version; OSVersion = $osStr } | ConvertTo-Json -Compress
    }
    exit 0
}

# ------------------------------------------------------------
# Download Mode
# ------------------------------------------------------------

if ($DryRun) { Write-Host " [DryRun Mode]" -ForegroundColor Yellow }
Write-Host " Found: $($product.ModelName)" -ForegroundColor Green
Write-Host "--- GetDriver - ASUS SCCM Pack Downloader ---" -ForegroundColor Magenta

# 5. Resolve Destination Root
if ([string]::IsNullOrWhiteSpace($Destination)) {
    if ($config -and $config.Shared_Drivers_Location) {
        $Destination = $config.Shared_Drivers_Location
    } else {
        Write-Error "Destination not specified and 'Shared_Drivers_Location' not defined in config.psd1. Please specify a destination root or define the variable."
        exit 1
    }
}

$ModelDest = Join-Path $Destination $product.ModelName

# 6. Handle Selection
$selectedPack = $null
$IsInteractive = [Environment]::UserInteractive -and -not $PSBoundParameters.ContainsKey('AutoExpand') -and -not $Version -and -not $PSBoundParameters.ContainsKey('OSVersion') -and -not $DryRun

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
    $filtered = $sccmPacks
    if ($Version) { $filtered = $filtered | Where-Object { $_.Version -eq $Version } }
    if ($PSBoundParameters.ContainsKey('OSVersion')) { $filtered = $filtered | Where-Object { $_.Title -match $OSVersion } }

    $selectedPack = $filtered | Select-Object -First 1
    if (-not $selectedPack) {
        Write-Error "No matching package found for Version='$Version' and OSVersion='$OSVersion'."
        exit 1
    }
}

# 7. Download
$fileName = (Split-Path $selectedPack.DownloadUrl -Leaf).Split('?')[0]
$outPath = Join-Path $ModelDest $fileName

Write-Host "`nDownloading: $fileName..." -ForegroundColor Cyan
Write-Host "Target: $outPath" -ForegroundColor DarkGray

if ($DryRun) {
    Write-Host "[DryRun] Would download from $($selectedPack.DownloadUrl) to $outPath" -ForegroundColor Yellow
} else {
    if (-not (Test-Path $ModelDest)) { New-Item -ItemType Directory -Path $ModelDest -Force | Out-Null }
    Invoke-WebRequest -Uri $selectedPack.DownloadUrl -OutFile $outPath
}

if ($DryRun -or (Test-Path $outPath)) {
    if (-not $DryRun) { Write-Host "Download complete." -ForegroundColor Green }

    # 8. Extraction
    $doExpand = $true
    if ($PSBoundParameters.ContainsKey('AutoExpand')) {
        if ($AutoExpand -is [switch] -or $AutoExpand -is [bool]) { $doExpand = [bool]$AutoExpand }
        else { [bool]::TryParse($AutoExpand, [ref]$doExpand) | Out-Null }
    }

    if ($doExpand) { Expand-DriverPackage -ZipPath $outPath -DestDir $ModelDest }
} else {
    Write-Error "Download failed."
}
