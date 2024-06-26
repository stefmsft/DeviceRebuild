rem Re format the partition
diskpart /s C:\Windows\Setup\Scripts\Add-RecoLetter.txt

:rem == Register the location of the recovery tools ==
Reagentc /Setreimage /Path R:\Recovery\WindowsRE /Target C:\Windows

:rem == Enable WinRE wim ==
Reagentc /enable /Target C:\Windows

:rem == Verify the configuration status of the images. ==
Reagentc /Info /Target C:\Windows
