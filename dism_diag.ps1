#Requires -RunAsAdministrator

Write-Host "=== DISM processes ===" -ForegroundColor Cyan
$dism = Get-Process -Name dism -ErrorAction SilentlyContinue
if ($dism) { $dism | Select-Object Id, CPU, WS, StartTime | Format-Table -Auto }
else { Write-Host "  (none running)" }

Write-Host ""
Write-Host "=== Mounted WIM images ===" -ForegroundColor Cyan
$mounted = Get-WindowsImage -Mounted
if ($mounted) {
    $mounted | Select-Object ImagePath, MountPath, MountMode, MountStatus | Format-List
} else {
    Write-Host "  (no images currently mounted)"
}

Write-Host ""
Write-Host "=== C:\WinPE_Mount top-level ===" -ForegroundColor Cyan
if (Test-Path "C:\WinPE_Mount") {
    Get-ChildItem "C:\WinPE_Mount" | Select-Object Name, LastWriteTime | Format-Table -Auto
} else {
    Write-Host "  Directory does not exist."
}
