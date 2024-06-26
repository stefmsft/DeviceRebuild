rem == CaptureImage.bat ==

rem == These commands capture Device Parttion

rem    Usage:   CaptureImage ModelTag 

rem Assign letters on all the partitions
diskpart /s Mount-AllLetters.txt

rem Add or remove the partition your wish

Dism /Capture-Image /ImageFile:%1-OS.wim /CaptureDir:W:\ /Name:%1-OS
Dism /Capture-Image /ImageFile:%1-SYSTEM.wim /CaptureDir:S:\ /Name:%1-SYSTEM
Dism /Capture-Image /ImageFile:%1-RECOVERY.wim /CaptureDir:R:\ /Name:%1-RECOVERY
Dism /Capture-Image /ImageFile:%1-MYASUS.wim /CaptureDir:M:\ /Name:%1-MYASUS

diskpart /s DiskVol.txt > GeometryDsk.txt
