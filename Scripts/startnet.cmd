wpeinit

@echo off
setlocal enabledelayedexpansion

:: Specify the file and directory to search for config file
set "targetFileIni=config.ini"
set "targetDirIni=\"

:: Specify the file and directory to search for the default script
set "targetScript=ApplyImage.bat"
set "targetDirScript=\"

:: Loop through all possible drive letters
for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
	if exist %%D:%targetDirIni%\%targetFileIni% (
		echo script found on drive %%D:
		@echo off
		%%D:
		cd %targetDirIni%
		:: Read the .ini file and set variables
		for /f "tokens=1,2 delims== " %%A in (%targetFileIni%) do (
    		set "%%A=%%B"
		)
	)
)

:: Verify the variables are set
echo ModelString=%ModelString%
echo targetScript=%targetScript%
pause
@echo off

:: Check if modelString is set
if not defined modelString (
    echo Error: modelString is not set or is empty.
    endlocal
    exit /B
)

:: Get the SystemProductName from the registry
for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\SystemInformation" /v SystemProductName 2^>nul') do (
    set "systemProductName=%%B"
)

:: Remove any surrounding quotes (if any)
set "systemProductName=%systemProductName:"=%"


:: Extract the relevant substring from systemProductName
set "substring=%systemProductName:~16,7%"

:: Check if the extracted substring matches the modelString
if /i "%substring%"=="%modelString%" (
	:: Loop through all possible drive letters
	for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
		if exist %%D:%targetDirScript%\%targetScript% (
			echo script found on drive %%D:
			@echo off
			%%D:
			cd %targetDirScript%
			%targetScript% "%modelString%"
			echo end of script ... Reboot in 5 seconds
			@echo off
			ping -n 6 127.0.0.1 >nul
			w:\Windows\system32\shutdown /s /t 0
			endlocal
			exit /B
		)
	)
	echo Apply.bat not found on any drive.
) else (
	echo SystemProductName does not match the targeted model.
)

endlocal