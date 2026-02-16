@echo off
setlocal enabledelayedexpansion

rem == ApplyImage.bat ==
rem == Deploys Windows image files to partitions and configures the system.
rem == Supports both standard WIM and split SWM files for OS partition.
rem == Usage: ApplyImage ModelTag

rem == Setup logging ==
rem Create timestamp for log file (handles various date/time formats)
set "LOGDATE=%date:~-4%%date:~-7,2%%date:~-10,2%"
set "LOGTIME=%time:~0,2%%time:~3,2%%time:~6,2%"
set "LOGTIME=%LOGTIME: =0%"
set "LOGFILE=ApplyImage_%LOGDATE%_%LOGTIME%.log"

rem Start logging - all output goes to both console and file
call :log "============================================================"
call :log "  ApplyImage - Windows Deployment Script"
call :log "  Model: %~1"
call :log "  Log file: %LOGFILE%"
call :log "============================================================"
call :log ""

rem == Validate parameter ==
if "%~1"=="" (
    call :log "ERROR: Model tag parameter is required."
    call :log "Usage: ApplyImage ModelTag"
    exit /b 1
)

set "MODEL=%~1"

rem == Set high-performance power scheme to speed deployment ==
call :log "[1/12] Setting high-performance power scheme..."
powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >> "%LOGFILE%" 2>&1

rem == Validate DiskNumber is set and is a number ==
if not defined DiskNumber (
    call :log "ERROR: DiskNumber not defined in config.ini"
    call :log "       Please set DiskNumber to the target disk number."
    call :log "       Use 'diskpart' and 'list disk' to find the correct number."
    pause
    exit /b 1
)

rem Check DiskNumber is not empty
if "%DiskNumber%"=="" (
    call :log "ERROR: DiskNumber is empty in config.ini"
    pause
    exit /b 1
)

call :log "      Target DiskNumber=%DiskNumber%"

rem ============================================================
rem == SAFETY CHECK: Prevent formatting the USB key
rem ============================================================
call :log ""
call :log "[2/12] Safety check - Verifying target disk..."

rem Get the drive letter where this script is running from
set "SCRIPTDRIVE=%~d0"
set "SCRIPTLETTER=%SCRIPTDRIVE:~0,1%"
call :log "      Script running from: %SCRIPTDRIVE%"

rem Check if DEPLOYKEY.marker exists on current drive (confirms this is USB key)
if exist "%SCRIPTDRIVE%\DEPLOYKEY.marker" (
    call :log "      USB deployment key confirmed: %SCRIPTDRIVE%"
) else (
    call :log "WARNING: DEPLOYKEY.marker not found on %SCRIPTDRIVE%"
    call :log "         Cannot verify this is the deployment USB key."
)

rem SAFETY CHECK: Check ALL partitions on target disk for marker
call :log "      Checking target disk partitions for safety marker..."
set "MARKER_FOUND=0"

rem Check partitions 1 through 4 (covers typical USB layouts)
for %%P in (1 2 3 4) do (
    if "!MARKER_FOUND!"=="0" (
        (
            echo select disk %DiskNumber%
            echo select partition %%P
            echo assign letter=Z
        ) > TempMount.txt
        diskpart /s TempMount.txt > nul 2>&1

        if exist "Z:\DEPLOYKEY.marker" (
            set "MARKER_FOUND=1"
            call :log "      DEPLOYKEY.marker found on partition %%P!"
        )

        (
            echo select volume Z
            echo remove letter=Z
        ) > TempUnmount.txt
        diskpart /s TempUnmount.txt > nul 2>&1
    )
)

rem Cleanup temp files
del TempMount.txt 2>nul
del TempUnmount.txt 2>nul

if "%MARKER_FOUND%"=="1" (
    call :log ""
    call :log "============================================================"
    call :log "  CRITICAL ERROR - DEPLOYMENT ABORTED"
    call :log "============================================================"
    call :log "  DEPLOYKEY.marker found on target Disk %DiskNumber%!"
    call :log "  This appears to be the USB deployment key, NOT the target device."
    call :log ""
    call :log "  The USB key may have been assigned as Disk %DiskNumber% by the BIOS."
    call :log "  Please verify disk numbers using: diskpart -> list disk"
    call :log ""
    call :log "  Update config.ini with the correct DiskNumber for the"
    call :log "  INTERNAL device disk, not the USB key."
    call :log "============================================================"
    pause
    exit /b 1
)

call :log "      Safety check passed - Target disk %DiskNumber% is not the USB key."

rem ============================================================
rem == PHASE 1: Partition the disk
rem ============================================================
call :log ""
call :log "[3/12] Creating partitions on Disk %DiskNumber%..."

rem Generate dynamic partition script with configured disk number
(
    echo select disk %DiskNumber%
    echo clean
    echo convert gpt
    echo rem == 1. System partition =========================
    echo create partition efi size=260
    echo format quick fs=fat32 label="System"
    echo assign letter="S"
    echo rem == 2. Microsoft Reserved ^(MSR^) partition =======
    echo create partition msr size=16
    echo rem == 3. Windows partition ========================
    echo create partition primary
    echo shrink minimum=2500
    echo format quick fs=ntfs label="Windows"
    echo assign letter="W"
    echo rem == 4. Recovery partition ======================
    echo create partition primary size=2000
    echo format quick fs=ntfs label="Recovery"
    echo assign letter="R"
    echo set id="de94bba4-06d1-4d40-a16a-bfd50179d6ac"
    echo gpt attributes=0x8000000000000001
    echo rem == 5. MyAsus partition ========================
    echo create partition primary
    echo format quick fs=ntfs label="MyASUS"
    echo assign letter="M"
    echo set id="ebd0a0a2-b9e5-4433-87c0-68b6b72699c7"
    echo gpt attributes=0x8000000000000000
    echo list volume
    echo exit
) > CreatePartitions-UEFI-Dynamic.txt

call :log "      Running diskpart..."
diskpart /s CreatePartitions-UEFI-Dynamic.txt >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "ERROR: DiskPart failed to create partitions."
    exit /b 1
)

rem ============================================================
rem == PHASE 2: Apply OS image (WIM or SWM)
rem ============================================================
call :log ""
call :log "[4/12] Applying OS image to W:\..."

rem Check for split WIM files first (SWM takes priority)
if exist "%MODEL%-OS.swm" (
    call :log "      Detected split WIM files (SWM), applying with /SWMFile..."
    dism /Apply-Image /ImageFile:"%MODEL%-OS.swm" /SWMFile:"%MODEL%-OS*.swm" /Index:1 /ApplyDir:W:\ >> "%LOGFILE%" 2>&1
) else if exist "%MODEL%-OS.wim" (
    call :log "      Applying standard WIM file..."
    dism /Apply-Image /ImageFile:"%MODEL%-OS.wim" /Index:1 /ApplyDir:W:\ >> "%LOGFILE%" 2>&1
) else (
    call :log "ERROR: No OS image found. Expected %MODEL%-OS.swm or %MODEL%-OS.wim"
    exit /b 1
)

if errorlevel 1 (
    call :log "ERROR: Failed to apply OS image."
    exit /b 1
)

rem ============================================================
rem == PHASE 3: Inject drivers (optional)
rem ============================================================
call :log ""
call :log "[5/12] Checking for drivers to inject..."

if exist "Drivers\%MODEL%" (
    call :log "      Found driver directory: Drivers\%MODEL%"
    call :log "      Injecting drivers into offline image..."
    dism /Image:W:\ /Add-Driver /Driver:"Drivers\%MODEL%" /Recurse >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: Some drivers may have failed to inject."
    ) else (
        call :log "      Drivers injected successfully."
    )
) else (
    call :log "      No driver directory found at Drivers\%MODEL%"
    call :log "      Skipping driver injection."
)

rem ============================================================
rem == PHASE 4: OOBE BypassNRO (skip network requirement)
rem ============================================================
call :log ""
if /i "%BypassNRO%"=="1" (
    call :log "[6/12] Applying OOBE BypassNRO (skip network requirement)..."
    reg load HKLM\OFFLINE "W:\Windows\System32\Config\SOFTWARE" >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: Failed to load offline registry hive."
    ) else (
        reg add "HKLM\OFFLINE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f >> "%LOGFILE%" 2>&1
        if errorlevel 1 (
            call :log "WARNING: Failed to set BypassNRO registry key."
        ) else (
            call :log "      BypassNRO registry key set successfully."
        )
        reg unload HKLM\OFFLINE >> "%LOGFILE%" 2>&1
    )
) else (
    call :log "[6/12] Skipping OOBE BypassNRO (disabled in config.ini)."
)

rem ============================================================
rem == PHASE 5: Configure boot files
rem ============================================================
call :log ""
call :log "[7/12] Configuring boot files on S:\..."
W:\Windows\System32\bcdboot W:\Windows /s S: >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "ERROR: Failed to configure boot files."
    exit /b 1
)

rem ============================================================
rem == PHASE 5: Set up Recovery partition (winre.wim)
rem ============================================================
call :log ""
call :log "[8/12] Setting up Recovery partition (winre.wim)..."

mkdir R:\Recovery\WindowsRE 2>nul
set "WINRE_READY=0"

rem Primary method: copy winre.wim from OS image + inject iRST drivers if available
if exist "W:\Windows\System32\Recovery\winre.wim" (
    rem Check if iRST drivers are needed and available
    set "RST_PATH="
    if exist "Drivers\%MODEL%\Rapid Storage" (
        set "RST_PATH=Drivers\%MODEL%\Rapid Storage"
    )

    if defined RST_PATH (
        call :log "      Copying winre.wim from OS image..."
        xcopy /H "W:\Windows\System32\Recovery\winre.wim" "R:\Recovery\WindowsRE\" >> "%LOGFILE%" 2>&1
        if errorlevel 1 (
            call :log "WARNING: Failed to copy winre.wim from OS image."
        ) else (
            call :log "      Injecting Rapid Storage drivers into winre.wim..."
            set "WINRE_MOUNT=W:\WinRE_Mount"
            mkdir "!WINRE_MOUNT!" 2>nul
            dism /Mount-Image /ImageFile:"R:\Recovery\WindowsRE\winre.wim" /Index:1 /MountDir:"!WINRE_MOUNT!" >> "%LOGFILE%" 2>&1
            if errorlevel 1 (
                call :log "WARNING: Failed to mount winre.wim for driver injection."
            ) else (
                dism /Image:"!WINRE_MOUNT!" /Add-Driver /Driver:"!RST_PATH!" /Recurse >> "%LOGFILE%" 2>&1
                if errorlevel 1 (
                    call :log "WARNING: Failed to inject Rapid Storage drivers into winre.wim."
                ) else (
                    call :log "      Rapid Storage drivers injected into winre.wim successfully."
                )
                dism /Unmount-Image /MountDir:"!WINRE_MOUNT!" /Commit >> "%LOGFILE%" 2>&1
                if errorlevel 1 (
                    call :log "WARNING: Failed to unmount winre.wim. Attempting discard..."
                    dism /Unmount-Image /MountDir:"!WINRE_MOUNT!" /Discard >> "%LOGFILE%" 2>&1
                )
            )
            rmdir /s /q "!WINRE_MOUNT!" 2>nul
            set "WINRE_READY=1"
        )
    ) else (
        rem No Rapid Storage drivers found - fall back to RECOVERY.wim if available
        rem (a captured RECOVERY.wim likely already has the required storage drivers)
        if exist "%MODEL%-RECOVERY.wim" (
            call :log "      No Rapid Storage drivers found in Drivers\%MODEL%"
            call :log "      Falling back to %MODEL%-RECOVERY.wim (likely has storage drivers)..."
            dism /Apply-Image /ImageFile:"%MODEL%-RECOVERY.wim" /Index:1 /ApplyDir:R:\ >> "%LOGFILE%" 2>&1
            if errorlevel 1 (
                call :log "WARNING: Failed to apply Recovery image."
            ) else (
                set "WINRE_READY=1"
            )
        ) else (
            rem No drivers needed or no fallback - use vanilla winre.wim
            call :log "      Copying winre.wim from OS image (no storage drivers to inject)..."
            xcopy /H "W:\Windows\System32\Recovery\winre.wim" "R:\Recovery\WindowsRE\" >> "%LOGFILE%" 2>&1
            if errorlevel 1 (
                call :log "WARNING: Failed to copy winre.wim from OS image."
            ) else (
                set "WINRE_READY=1"
            )
        )
    )
) else if exist "%MODEL%-RECOVERY.wim" (
    rem No winre.wim in OS image - use captured RECOVERY.wim
    call :log "      winre.wim not found in OS image, using %MODEL%-RECOVERY.wim..."
    dism /Apply-Image /ImageFile:"%MODEL%-RECOVERY.wim" /Index:1 /ApplyDir:R:\ >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: Failed to apply Recovery image."
    ) else (
        set "WINRE_READY=1"
    )
) else (
    call :log "WARNING: No recovery source found."
    call :log "         Checked W:\Windows\System32\Recovery\winre.wim"
    call :log "         Checked %MODEL%-RECOVERY.wim"
)

rem Verify winre.wim is in place
if exist "R:\Recovery\WindowsRE\winre.wim" (
    call :log "      Verified: winre.wim present at R:\Recovery\WindowsRE\"
) else (
    call :log "ERROR: winre.wim not found at R:\Recovery\WindowsRE\"
    call :log "       Windows Recovery and Reset will NOT work."
)

rem ============================================================
rem == PHASE 6: Apply MyASUS partition
rem ============================================================
call :log ""
call :log "[9/12] Applying MyASUS image to M:\..."
if exist "%MODEL%-MYASUS.wim" (
    dism /Apply-Image /ImageFile:"%MODEL%-MYASUS.wim" /Index:1 /ApplyDir:M:\ >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: Failed to apply MyASUS image."
    )
) else (
    call :log "WARNING: MyASUS image not found: %MODEL%-MYASUS.wim"
)

rem ============================================================
rem == PHASE 7: Configure Push Button Reset (ResetConfig.xml)
rem ============================================================
call :log ""
call :log "[10/12] Configuring Push Button Reset..."

rem Create Recovery\OEM folder with proper permissions
call :log "      Creating W:\Recovery\OEM structure..."
mkdir W:\Recovery\OEM 2>nul
icacls W:\Recovery /inheritance:r >> "%LOGFILE%" 2>&1
icacls W:\Recovery /grant:r SYSTEM:(OI)(CI)(F) >> "%LOGFILE%" 2>&1
icacls W:\Recovery /grant:r *S-1-5-32-544:(OI)(CI)(F) >> "%LOGFILE%" 2>&1
takeown /f W:\Recovery /a >> "%LOGFILE%" 2>&1
attrib +H W:\Recovery >> "%LOGFILE%" 2>&1

rem Create ResetConfig.xml for non-standard partition layout
call :log "      Creating ResetConfig.xml..."
(
    echo ^<?xml version="1.0" encoding="utf-8"?^>
    echo ^<Reset^>
    echo   ^<SystemDisk^>
    echo     ^<MinSize^>75000^</MinSize^>
    echo     ^<DiskpartScriptPath^>ReCreatePartitions-UEFI.txt^</DiskpartScriptPath^>
    echo     ^<OSPartition^>3^</OSPartition^>
    echo     ^<WindowsREPartition^>4^</WindowsREPartition^>
    echo     ^<WindowsREPath^>Recovery\WindowsRE^</WindowsREPath^>
    echo     ^<Compact^>False^</Compact^>
    echo   ^</SystemDisk^>
    echo ^</Reset^>
) > W:\Recovery\OEM\ResetConfig.xml

rem Create bare-metal recovery DiskPart script (no select disk / no clean)
call :log "      Creating ReCreatePartitions-UEFI.txt..."
(
    echo select disk %DiskNumber%
    echo clean
    echo convert gpt
    echo create partition efi size=260
    echo format quick fs=fat32 label="System"
    echo assign letter="S"
    echo create partition msr size=16
    echo create partition primary
    echo shrink minimum=2500
    echo format quick fs=ntfs label="Windows"
    echo assign letter="W"
    echo create partition primary size=2000
    echo format quick fs=ntfs label="Recovery"
    echo assign letter="R"
    echo set id="de94bba4-06d1-4d40-a16a-bfd50179d6ac"
    echo gpt attributes=0x8000000000000001
    echo create partition primary
    echo format quick fs=ntfs label="MyASUS"
    echo assign letter="M"
    echo set id="ebd0a0a2-b9e5-4433-87c0-68b6b72699c7"
    echo gpt attributes=0x8000000000000000
    echo exit
) > W:\Recovery\OEM\ReCreatePartitions-UEFI.txt

call :log "      Push Button Reset configured."

rem ============================================================
rem == PHASE 8: Configure Windows Recovery Environment
rem ============================================================
call :log ""
call :log "[11/12] Configuring Windows Recovery Environment..."

rem Register the recovery image location (offline)
call :log "      Registering recovery image location..."
W:\Windows\System32\ReAgentC.exe /setreimage /path "R:\Recovery\WindowsRE" /target "W:\Windows" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :log "WARNING: Failed to register recovery image location."
)

rem Enable Windows Recovery Environment (requires /osguid in WinPE)
call :log "      Retrieving OS GUID from BCD store..."
set "OSGUID="
for /f "tokens=2 delims={}" %%G in ('bcdedit /store S:\EFI\Microsoft\Boot\BCD /enum {default} /v') do (
    if not defined OSGUID set "OSGUID={%%G}"
)

if defined OSGUID (
    call :log "      Found OS GUID: %OSGUID%"
    call :log "      Enabling Windows Recovery Environment..."
    W:\Windows\System32\ReAgentC.exe /enable /osguid %OSGUID% >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "WARNING: Failed to enable WinRE."
    )
) else (
    call :log "WARNING: Could not retrieve OS GUID from BCD store."
    call :log "         WinRE must be enabled manually after first boot."
)

rem ============================================================
rem == PHASE 9: Final verification
rem ============================================================
call :log ""
call :log "[12/12] Final verification..."

rem Display recovery configuration
call :log "      Recovery configuration status:"
W:\Windows\System32\ReAgentC.exe /info /target "W:\Windows" >> "%LOGFILE%" 2>&1

rem Verify ResetConfig.xml
if exist "W:\Recovery\OEM\ResetConfig.xml" (
    call :log "      ResetConfig.xml: Present"
) else (
    call :log "WARNING: ResetConfig.xml not found - Push Button Reset may fail."
)

rem Verify winre.wim
if exist "R:\Recovery\WindowsRE\winre.wim" (
    call :log "      winre.wim: Present"
) else (
    call :log "ERROR: winre.wim missing - Recovery will NOT work."
)

call :log ""
call :log "============================================================"
call :log "  Deployment completed successfully for model: %MODEL%"
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