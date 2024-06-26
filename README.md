# DeviceRebuild

# Purpose

Automatically rebuild a Windows Device based on a OOBE WIM capture of the model.

# Usage

- Plug the USB Key holding the rebuild content in the device.
- Reboot the device on the key
	- Reboot + Esc or Reboot + F2 + F8 on Asus devices
	- Then select the entry representing the first partition of the USB drive
- Wait for the end of the operation and the reboot of the device.
- If everything works correctly you'll end up in the OOBE for a fresh setup of the device.

>[!Warning] 
>If the model name of device doesn't match the ModelString variable define in the file config.ini in the root of the USB Drive, the operation will abort. This reduce the danger of erasing a device by accident.
>Nevertheless you should be careful before booting on the key. If this is a compatible model, this will lead in a complete wipe of the device without any warning

# To produce the USB Key

## Prerequisits

- You need to use an USB key big enough to hold the WIMs (16/32 Go should be ok)
- You need to have a WIM file containing a WinPE with a modified startnet.cmd.
>[!Nota] 
>The code of startnet.cmd used is in the git repo. refer to the [creating / Modifying startnet.cmd] section for inserting / modifying this code. 
- You need to prepare the WIM for a model
	- For that you reset a targeted model device so it boot in OOBE (Thru a "System reset" or by "sysprep /generalize /oobe / reboot")
	- From here you boot in WinPE and capture de partitions following the naming _ModelName_-PartitionName.WIM (_ModelName_ on 7 Characters, PartitionName part of the following list : OS,SYSTEM,RECOVERY,MYASUS)
>[!Nota] 
>There's a script (CaptureImage.bat) on the key (in the root directory) that can do the capture automatically for you. For that you need to modify the variable targetScript in config.ini file with the value "CaptureImage.bat". Then you boot on PE and the capture will be done automatically with the correct naming. See [WIM Repo] section for more details
- You need to sync the git repo containing the tools I produce.
	- They contain : 
		- A PowerShell script named ProduceKey.ps1
        - A variable file container named config.psd1
		- A Sub-Directory named Scripts containing the script that will be used to wipe the device

# USB Key Generation

When all the prerequisites are met, and the repo scripts gathered you plug an USB Key in your device.
Then you edit config.psd1 to reflect where is located the WinPE.WIM file (WinPE_WIM_Location variable) and the root directory holding the model WIMs (Device_Root_WIM_Location)
The script will abort is the content of those variable is not valid.
Launch the ProduceKey powershell script and follow the questions

Formating the USB key is optionnal. If the option is chosen then the copy of the PE content is automatically done. Otherwise you can choose to skip the operation as well
In all case the scripts are refreshed to the key and you are finally asked if you wish the script to copy over the WIMs for a model.

See below for more details

```Shell
.\ProduceKey.ps1
Do you want to format a new drive ? Y/N : Y
USB storage devices with media found:                                                                                      Disk 1: SanDisk Cruzer Fit - Partition 1 Size: 2GB - Drive Letter: P                                                       Please enter a Drive Letter or press 'Escape' to exit:                                                                     The drive letter P: is valid, not a system drive, and does not exceed 64 GB.                                               The disk number for drive letter P: is: 1                                                                                  Are you sure you want to clean Disk  This will erase all data. Type 'Y' to confirm : Y
Running diskpart with PrepareUSB.txt...

Microsoft DiskPart version 10.0.22621.1

Copyright (C) Microsoft Corporation.
On computer: OF2300479-NB

Disk 1 is now the selected disk.

  Disk ###  Status         Size     Free     Dyn  Gpt
  --------  -------------  -------  -------  ---  ---
  Disk 0    Online          953 GB  1024 KB        *
* Disk 1    Online           29 GB    27 GB
  Disk 2    No Media           0 B      0 B

Disk attributes cleared successfully.

DiskPart succeeded in cleaning the disk.

DiskPart successfully converted the selected disk to MBR format.

DiskPart succeeded in creating the specified partition.

DiskPart marked the current partition as active.

  100 percent completed

DiskPart successfully formatted the volume.

DiskPart successfully assigned the drive letter or mount point.

DiskPart succeeded in creating the specified partition.

  100 percent completed

DiskPart successfully formatted the volume.

DiskPart successfully assigned the drive letter or mount point.

  Partition ###  Type              Size     Offset
  -------------  ----------------  -------  -------
  Partition 1    Primary           2048 MB  1024 KB
* Partition 2    Primary             27 GB  2049 MB

Leaving DiskPart...
diskpart has completed.
Apply PE on the P: Partition
Please select the destination drive for scripts and WIMs
USB storage devices with media found:
Disk 1: SanDisk Cruzer Fit - Partition 1 Size: 2GB - Drive Letter: P
Disk 1: SanDisk Cruzer Fit - Partition 2 Size: 27.25GB - Drive Letter: I
Please enter a Drive Letter or press 'Escape' to exit:
The drive letter I: is valid, not a system drive, and does not exceed 64 GB.
Copy the Script files to I:
Do you want also to provide a set of model WIMs on drive I: ? Y/N : Y
Copy the model WIMs to the key
1: B9400CB
2: B9450FA
Please choose a number between 1 and 2: 2
Model B9450FA chosen
```

# Model WIM repo

## Structure

Pointed by the variable Device_Root_WIM_Location is the root directory where you keep the different captured done. The WIM file should be on a subdirectories representing the modelname on 7 characters (non less no more).
On the output above there were 2 directory available holding the name B9400CB and B9450FA

## Operation
Create a directory
Use the CaptureImage.bat script on a prepared device to generate the WIM on the key. Copy those WIMs inside a subdirectory holding the modelname for future deploiement.

# WinPE WIM

I officially won't provide the WinPE content in my repo.

You can use the ADK and follow the plethora of content explain how to produce a WinPE
This way you can custom it with your language or specific Drivers or Package you wish to add.
When its done you need to customize it with the startnet.cmd file provided in the repo.

For that follow the next section.

Finally capture it with a command like :

```Shell
Dism /Capture-Image /ImageFile:C:\Temp\PEMakeUSB.wim /CaptureDir:P:\ /Name:PEMakeUSB
```

# creating / Modifying startnet.cmd

To modify startnet.cmd do the following commands

```Shell
Dism /Mount-Image /ImageFile:"E:\sources\boot.wim" /index:1 /MountDir:"C:\WinPE_amd64\mount"
# Edit or copy the content from the git repo
notepad c:\WinPE_amd64\mount\Windows\System32\startnet.cmd
Dism /Unmount-Image /MountDir:"C:\WinPE_amd64\mount" /commit
```