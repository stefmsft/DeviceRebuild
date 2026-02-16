wpeinit

@echo off
setlocal enabledelayedexpansion

:: == Setup logging ==
:: Create timestamp for log file
set "LOGDATE=%date:~-4%%date:~-7,2%%date:~-10,2%"
set "LOGTIME=%time:~0,2%%time:~3,2%%time:~6,2%"
set "LOGTIME=%LOGTIME: =0%"
set "LOGFILE=Startnet_%LOGDATE%_%LOGTIME%.log"
set "USBDRIVE="

:: Specify the file and directory to search for config file
set "targetFileIni=config.ini"
set "targetDirIni=\"

:: Specify the file and directory to search for the default script
set "targetScript=ApplyImage.bat"
set "targetDirScript=\"

echo ============================================================
echo   Startnet - WinPE Boot Script
echo   Searching for config.ini...
echo ============================================================

:: Loop through all possible drive letters to find config.ini
for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
	if exist %%D:%targetDirIni%\%targetFileIni% (
		set "USBDRIVE=%%D:"
		echo Found config.ini on drive %%D:
		%%D:
		cd %targetDirIni%
		:: Read the .ini file and set variables
		for /f "tokens=1,2 delims== " %%A in (%targetFileIni%) do (
    		set "%%A=%%B"
		)
	)
)

:: If no USB drive found, log to X: (WinPE RAM drive)
if not defined USBDRIVE (
    set "LOGFILE=X:\%LOGFILE%"
    call :log "ERROR: config.ini not found on any drive"
    call :log "Searched all drives A-Z for %targetFileIni%"
    pause
    endlocal
    exit /B
)

:: Now we can start logging to USB drive
set "LOGFILE=%USBDRIVE%\%LOGFILE%"
call :log "============================================================"
call :log "  Startnet - WinPE Boot Script"
call :log "  Log file: %LOGFILE%"
call :log "============================================================"
call :log ""
call :log "USB drive found: %USBDRIVE%"
call :log ""

:: Log the variables read from config.ini
call :log "Configuration from config.ini:"
call :log "  ModelString=%ModelString%"
call :log "  targetScript=%targetScript%"
call :log "  DiskNumber=%DiskNumber%"
call :log "  BypassNRO=%BypassNRO%"
call :log ""

:: Check if modelString is set
if not defined ModelString (
    call :log "ERROR: ModelString is not set or is empty in config.ini"
    pause
    endlocal
    exit /B
)

:: Get the SystemProductName from the registry
call :log "Reading SystemProductName from registry..."
for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\SystemInformation" /v SystemProductName 2^>nul') do (
    set "systemProductName=%%B"
)

:: Remove any surrounding quotes (if any)
set "systemProductName=%systemProductName:"=%"
call :log "  SystemProductName: %systemProductName%"
call :log ""

:: Check if modelString is contained anywhere in systemProductName
:: Using string substitution: if replacing modelString changes the string, it was found
call :log "Checking if ModelString is present in SystemProductName..."
call set "checkString=%%systemProductName:%ModelString%=%%"

if /i not "%checkString%"=="%systemProductName%" (
	:: Model string found in SystemProductName - proceed with deployment
	call :log "SUCCESS: Model match found!"
	call :log "  '%ModelString%' is present in '%systemProductName%'"
	call :log ""

	:: Loop through all possible drive letters to find target script
	call :log "Searching for %targetScript%..."
	for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
		if exist %%D:%targetDirScript%\%targetScript% (
			call :log "Found %targetScript% on drive %%D:"
			call :log ""
			call :log "============================================================"
			call :log "  READY TO EXECUTE"
			call :log "============================================================"
			call :log "  Script: %targetScript%"
			call :log "  Model: %ModelString%"
			call :log "  Target Disk: %DiskNumber%"
			call :log "============================================================"
			call :log ""
			set /p CONFIRM="Continue with deployment? [Y/N]: "
			if /i not "!CONFIRM!"=="Y" (
				call :log "Deployment aborted by user."
				endlocal
				exit /B
			)
			call :log ""
			call :log "Launching %targetScript%..."
			call :log "============================================================"
			%%D:
			cd %targetDirScript%
			set "DiskNumber=%DiskNumber%"
			%targetScript% "%ModelString%"
			call :log "============================================================"
			call :log "%targetScript% completed."
			call :log "System will shut down in 5 seconds..."
			ping -n 6 127.0.0.1 >nul
			w:\Windows\system32\shutdown /s /t 0
			endlocal
			exit /B
		)
	)
	call :log "ERROR: %targetScript% not found on any drive."
) else (
	call :log ""
	call :log "============================================================"
	call :log "  MODEL MISMATCH - Deployment aborted for safety"
	call :log "============================================================"
	call :log "  Expected model string: '%ModelString%'"
	call :log "  Device SystemProductName: '%systemProductName%'"
	call :log "  checkString result: '%checkString%'"
	call :log "============================================================"
	call :log ""
)

call :log "Script ended. Check log file: %LOGFILE%"
pause
endlocal
exit /B

:: ============================================================
:: == SUBROUTINE: Log message to console and file
:: ============================================================
:log
echo(%~1
echo(%~1 >> "%LOGFILE%"
goto :eof
