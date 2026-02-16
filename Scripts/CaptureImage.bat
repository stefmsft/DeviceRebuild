@echo off
setlocal enabledelayedexpansion

rem == CaptureImage.bat ==
rem == Captures device partitions to WIM files
rem == Usage: CaptureImage ModelTag

rem == Setup logging ==
set "LOGDATE=%date:~-4%%date:~-7,2%%date:~-10,2%"
set "LOGTIME=%time:~0,2%%time:~3,2%%time:~6,2%"
set "LOGTIME=%LOGTIME: =0%"
set "LOGFILE=CaptureImage_%LOGDATE%_%LOGTIME%.log"

call :log "============================================================"
call :log "  CaptureImage - Partition Capture Script"
call :log "  Model: %~1"
call :log "  Log file: %LOGFILE%"
call :log "============================================================"
call :log ""

rem == Validate parameter ==
if "%~1"=="" (
    call :log "ERROR: Model tag parameter is required."
    call :log "Usage: CaptureImage ModelTag"
    exit /b 1
)

set "MODEL=%~1"

rem == Set high-performance power scheme to speed capture ==
call :log "[1/6] Setting high-performance power scheme..."
powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >> "%LOGFILE%" 2>&1

rem ============================================================
rem == PHASE 1: Assign drive letters to partitions
rem ============================================================
call :log ""
call :log "[2/6] Assigning drive letters to partitions..."

if exist "Mount-AllLetters.txt" (
    diskpart /s Mount-AllLetters.txt >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: DiskPart may have encountered issues."
    )
) else (
    call :log "WARNING: Mount-AllLetters.txt not found. Assuming letters already assigned."
)

rem ============================================================
rem == PHASE 2: Capture OS partition
rem ============================================================
call :log ""
call :log "[3/6] Capturing OS partition (W:\) to %MODEL%-OS.wim..."
call :log "      This may take several minutes..."

dism /Capture-Image /ImageFile:"%MODEL%-OS.wim" /CaptureDir:W:\ /Name:"%MODEL%-OS" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "ERROR: Failed to capture OS partition."
) else (
    call :log "      OS partition captured successfully."
)

rem ============================================================
rem == PHASE 3: Capture SYSTEM partition
rem ============================================================
call :log ""
call :log "[4/6] Capturing SYSTEM partition (S:\) to %MODEL%-SYSTEM.wim..."

dism /Capture-Image /ImageFile:"%MODEL%-SYSTEM.wim" /CaptureDir:S:\ /Name:"%MODEL%-SYSTEM" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "WARNING: Failed to capture SYSTEM partition."
) else (
    call :log "      SYSTEM partition captured successfully."
)

rem ============================================================
rem == PHASE 4: Capture RECOVERY partition
rem ============================================================
call :log ""
call :log "[5/6] Capturing RECOVERY partition (R:\) to %MODEL%-RECOVERY.wim..."

dism /Capture-Image /ImageFile:"%MODEL%-RECOVERY.wim" /CaptureDir:R:\ /Name:"%MODEL%-RECOVERY" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "WARNING: Failed to capture RECOVERY partition."
) else (
    call :log "      RECOVERY partition captured successfully."
)

rem ============================================================
rem == PHASE 5: Capture MYASUS partition
rem ============================================================
call :log ""
call :log "[6/6] Capturing MYASUS partition (M:\) to %MODEL%-MYASUS.wim..."

dism /Capture-Image /ImageFile:"%MODEL%-MYASUS.wim" /CaptureDir:M:\ /Name:"%MODEL%-MYASUS" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "WARNING: Failed to capture MYASUS partition (may not exist on this device)."
) else (
    call :log "      MYASUS partition captured successfully."
)

rem ============================================================
rem == Save disk geometry information
rem ============================================================
call :log ""
call :log "Saving disk geometry information..."

if exist "DiskVol.txt" (
    diskpart /s DiskVol.txt > GeometryDsk.txt 2>&1
    call :log "      Disk geometry saved to GeometryDsk.txt"
) else (
    call :log "      DiskVol.txt not found, skipping geometry capture."
)

rem ============================================================
rem == List captured files
rem ============================================================
call :log ""
call :log "Captured WIM files:"
for %%F in ("%MODEL%-*.wim") do (
    call :log "      %%F"
)

call :log ""
call :log "============================================================"
call :log "  Capture completed for model: %MODEL%"
call :log "  Log saved to: %LOGFILE%"
call :log "============================================================"

endlocal
exit /b 0

rem ============================================================
rem == SUBROUTINE: Log message to console and file
rem ============================================================
:log
echo(%~1
echo(%~1 >> "%LOGFILE%"
goto :eof
