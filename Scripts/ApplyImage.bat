rem == ApplyImage.bat ==

rem == These commands deploy a specified Windows
rem    image file to the Windows partition, and configure
rem    the system partition.

rem    Usage:   ApplyImage ModelTag 

rem == Set high-performance power scheme to speed deployment ==
call powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

rem Re format the partition
diskpart /s CreatePartitions-UEFI.txt

rem == Apply the image to the Windows partition ==
dism /Apply-Image /ImageFile:%1-OS.wim /Index:1 /ApplyDir:W:\

rem == Copy boot files to the System partition ==
W:\Windows\System32\bcdboot W:\Windows /s S:

rem == Apply the image to the Recovery partition ==
dism /Apply-Image /ImageFile:%1-RECOVERY.wim /Index:1 /ApplyDir:R:\

rem == Apply the image to the Recovery partition ==
dism /Apply-Image /ImageFile:%1-MYASUS.wim /Index:1 /ApplyDir:M:\

:rem == Verify the configuration status of the images. ==
W:\Windows\System32\Reagentc /Info /Target W:\Windows

mkdir W:\Windows\Setup\Scripts
icacls W:\Windows\Setup\Scripts /grant SYSTEM:(OI)(CI)(F)

copy PostSetupScript.bat W:\Windows\Setup\Scripts\PostSetupScript.bat
icacls W:\Windows\Setup\Scripts\PostSetupScript.bat /grant SYSTEM:(F)

copy Add-RecoLetter.txt W:\Windows\Setup\Scripts\Add-RecoLetter.txt
icacls W:\Windows\Setup\Scripts\Add-RecoLetter.txt /grant SYSTEM:(F)

echo apply unattend
dism /image:W:\ /apply-unattend:unattend.xml
copy /Y unattend.xml W:\unattend.xml
rem pause

@echo off