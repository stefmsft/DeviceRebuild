@echo off
setlocal

rem == UpdateStartnet.bat ==
rem == Updates startnet.cmd inside the WinPE boot.wim on a USB key
rem == Usage: UpdateStartnet.bat DriveLetter
rem == Example: UpdateStartnet.bat P

echo ============================================================
echo   UpdateStartnet - Update WinPE startnet.cmd
echo ============================================================
echo.

rem == Validate parameter ==
if "%~1"=="" (
    echo ERROR: Drive letter parameter is required.
    echo Usage: UpdateStartnet.bat DriveLetter
    echo Example: UpdateStartnet.bat P
    exit /b 1
)

rem == Configuration ==
set "MOUNTDIR=C:\Wim_Mount"

rem == Derived paths ==
set "DRIVE=%~1:"
set "BOOTWIN=%DRIVE%\sources\boot.wim"
set "SCRIPTDIR=%~dp0"

rem == Check if boot.wim exists ==
if not exist "%BOOTWIN%" (
    echo ERROR: boot.wim not found at %BOOTWIN%
    echo        Make sure %DRIVE% is the WinPE partition.
    exit /b 1
)

rem == Check if new startnet.cmd exists ==
if not exist "%SCRIPTDIR%Scripts\startnet.cmd" (
    echo ERROR: startnet.cmd not found in %SCRIPTDIR%Scripts\
    exit /b 1
)

rem == Create mount directory ==
echo [1/4] Creating mount directory...
if exist "%MOUNTDIR%" (
    echo       Mount directory already exists, cleaning up...
    rmdir /s /q "%MOUNTDIR%" 2>nul
)
mkdir "%MOUNTDIR%"
if errorlevel 1 (
    echo ERROR: Failed to create mount directory.
    exit /b 1
)

rem == Mount the boot.wim ==
echo [2/4] Mounting boot.wim from %DRIVE%...
dism /Mount-Image /ImageFile:"%BOOTWIN%" /index:1 /MountDir:"%MOUNTDIR%"
if errorlevel 1 (
    echo ERROR: Failed to mount boot.wim
    rmdir /s /q "%MOUNTDIR%" 2>nul
    exit /b 1
)

rem == Copy the new startnet.cmd ==
echo [3/4] Copying new startnet.cmd...
copy /Y "%SCRIPTDIR%Scripts\startnet.cmd" "%MOUNTDIR%\Windows\System32\startnet.cmd"
if errorlevel 1 (
    echo ERROR: Failed to copy startnet.cmd
    echo        Unmounting without saving changes...
    dism /Unmount-Image /MountDir:"%MOUNTDIR%" /discard
    rmdir /s /q "%MOUNTDIR%" 2>nul
    exit /b 1
)

rem == Unmount and commit ==
echo [4/4] Unmounting and committing changes...
dism /Unmount-Image /MountDir:"%MOUNTDIR%" /commit
if errorlevel 1 (
    echo ERROR: Failed to unmount boot.wim
    echo        Trying to discard and cleanup...
    dism /Unmount-Image /MountDir:"%MOUNTDIR%" /discard 2>nul
    rmdir /s /q "%MOUNTDIR%" 2>nul
    exit /b 1
)

rem == Cleanup ==
rmdir /s /q "%MOUNTDIR%" 2>nul

echo.
echo ============================================================
echo   startnet.cmd updated successfully on %DRIVE%
echo ============================================================

endlocal
exit /b 0
