#Requires -RunAsAdministrator
<#
.SYNOPSIS
    RefreshDrivers.ps1 - Automates driver updates for the shared driver library.

.DESCRIPTION
    Checks for driver updates for models defined in config.psd1 or found in the 
    Shared_Drivers_Location. Infers local versions from SCCM ReleaseNote XLSX files.
    Handles safe updates via a temporary 'olddrv' folder.

.PARAMETER Log
    Enable logging to file.

.PARAMETER DryRun
    If specified, performs all logic (version checking, comparison) but skips 
    the actual download, folder movement, and extraction.
#>
param(
    [switch]$Log,
    [switch]$DryRun
)

# ============================================================
# Module + logging init
# ============================================================
Import-Module (Join-Path $PSScriptRoot "DeviceRebuild.psm1") -Force

$Script:LogEnabled = $Log.IsPresent
$Script:LogFile    = if ($Script:LogEnabled) { "RefreshDrivers_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" } else { $null }
Initialize-Logging -Enabled $Script:LogEnabled -LogFile $Script:LogFile

if ($DryRun) { Write-Banner "RefreshDrivers - [DRY RUN MODE]" }
else { Write-Banner "RefreshDrivers - Driver Library Maintenance" }

# ============================================================
# Load Configuration
# ============================================================
$ConfigPath = Join-Path $PSScriptRoot "config.psd1"
if (-not (Test-Path $ConfigPath)) {
    Write-Log "config.psd1 not found. Aborting." -Level Error
    exit 1
}

try {
    if (Get-Command Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
        $config = Import-PowerShellDataFile -Path $ConfigPath
    } else {
        $config = Invoke-Expression (Get-Content $ConfigPath -Raw)
    }
} catch {
    Write-Log "Failed to parse config.psd1: $_" -Level Error
    exit 1
}

$SharedRoot = $config.Shared_Drivers_Location
if (-not $SharedRoot -or -not (Test-Path $SharedRoot)) {
    Write-Log "Shared_Drivers_Location not defined or invalid: $SharedRoot" -Level Error
    exit 1
}

$DefaultOS = if ($config.Default_OS_Version) { $config.Default_OS_Version } else { "25H2" }
$TargetModels = if ($config.Target_Models) { $config.Target_Models } else { @() }

# ============================================================
# Helper Functions
# ============================================================

function Get-LocalVersion {
    param([string]$ModelDir)
    
    $xlsx = Get-ChildItem -Path $ModelDir -Filter "*_SCCM_ReleaseNote.xlsx" | Select-Object -First 1
    if (-not $xlsx) { return $null }

    if ($xlsx.Name -match '_(\d+\.\d+)_SCCM_') {
        return $Matches[1]
    }
    return $null
}

function Compare-Versions {
    param([string]$v1, [string]$v2)
    $n1 = [double]$v1
    $n2 = [double]$v2
    if ($n2 -gt $n1) { return 1 }
    if ($n1 -gt $n2) { return -1 }
    return 0
}

# ============================================================
# Discovery Phase
# ============================================================

$ModelsToProcess = @()

if ($TargetModels.Count -gt 0) {
    $ModelsToProcess = $TargetModels
} else {
    Write-Log "No Target_Models defined. Scanning $SharedRoot..."
    $ModelsToProcess = Get-ChildItem -Path $SharedRoot -Directory | Select-Object -ExpandProperty Name
}

Write-Log "Found $($ModelsToProcess.Count) models to check."

# ============================================================
# Processing Phase
# ============================================================

foreach ($Model in $ModelsToProcess) {
    Write-Section "Model: $Model"
    $ModelDir = Join-Path $SharedRoot $Model
    
    # 1. Check if update is available
    Write-Log "Checking latest version for $Model ($DefaultOS)..."
    $latestJson = & "$PSScriptRoot\GetDriver.ps1" -Model $Model -OSVersion $DefaultOS -GetLatestVersion 2>&1
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($latestJson)) {
        Write-Log "Could not fetch info for $Model. Error: $latestJson" -Level Warning
        continue
    }

    # If $latestJson contains multiple lines (e.g. searching for model...), we want the last one (JSON)
    $jsonOnly = ($latestJson -split "`n")[-1].Trim()
    
    try {
        $latestInfo = $jsonOnly | ConvertFrom-Json
    } catch {
        Write-Log "Failed to parse JSON for $Model. Raw: $jsonOnly" -Level Warning
        continue
    }
    
    $remoteVer = $latestInfo.Version
    $remoteOS  = $latestInfo.OSVersion

    # 2. Determine Action
    $NeedsDownload = $false
    $IsNewModel = -not (Test-Path $ModelDir)
    
    if ($IsNewModel) {
        Write-Log "Model directory missing. Preparing full download." -Level Info
        $NeedsDownload = $true
    } else {
        $localVer = Get-LocalVersion -ModelDir $ModelDir
        if (-not $localVer) {
            Write-Log "No SCCM ReleaseNote found in $ModelDir. Skipping (Nothing to update)." -Level Info
            continue
        }

        Write-Log "Local Version: $localVer | Remote Version: $remoteVer"
        if ((Compare-Versions -v1 $localVer -v2 $remoteVer) -eq 1) {
            Write-Log "Update available: $localVer -> $remoteVer" -Level Success
            $NeedsDownload = $true
        } else {
            Write-Log "Up to date." -Level Info
        }
    }

    # 3. Perform Download/Update
    if ($NeedsDownload) {
        if ($DryRun) {
            Write-Log "[DryRun] Would update $Model to version $remoteVer" -Level Yellow
            continue
        }

        $OldDrv = Join-Path $ModelDir "olddrv"
        $BackupCreated = $false

        if (-not $IsNewModel) {
            Write-Log "Backing up existing drivers to $OldDrv..."
            if (Test-Path $OldDrv) { Remove-Item $OldDrv -Recurse -Force }
            New-Item -ItemType Directory -Path $OldDrv -Force | Out-Null
            
            Get-ChildItem -Path $ModelDir -Exclude "olddrv" | ForEach-Object {
                Move-Item $_.FullName -Destination $OldDrv -Force
            }
            $BackupCreated = $true
        }

        Write-Log "Downloading and expanding version $remoteVer..."
        & "$PSScriptRoot\GetDriver.ps1" -Model $Model -Version $remoteVer -OSVersion $remoteOS -Destination $SharedRoot -AutoExpand $true
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Update successful." -Level Success
            if ($BackupCreated) {
                Write-Log "Cleaning up old drivers..."
                Remove-Item $OldDrv -Recurse -Force
            }
        } else {
            Write-Log "Update failed." -Level Error
            if ($BackupCreated) {
                Write-Log "Restoring old drivers..."
                Get-ChildItem -Path $OldDrv | ForEach-Object {
                    Move-Item $_.FullName -Destination $ModelDir -Force
                }
                Remove-Item $OldDrv -Recurse -Force
            }
            if ($IsNewModel) {
                if (Test-Path $ModelDir) { Remove-Item $ModelDir -Recurse -Force }
            }
        }
    }
}

if ($DryRun) { Write-Banner "RefreshDrivers - [DRY RUN COMPLETED]" }
else { Write-Banner "RefreshDrivers - Completed" }
