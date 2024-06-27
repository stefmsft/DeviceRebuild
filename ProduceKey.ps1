# Function to check if a drive has a Windows directory
function Has-WindowsDirectory {
    param (
        [string]$DriveLetter
    )

    # Check if the Windows directory exists on the drive
    return Test-Path "$DriveLetter\Windows"
}

# Function to check if a drive exceeds the size limit
function Is-DriveSizeExceeds {
    param (
        [string]$DriveLetter,
        [int64]$SizeLimit
    )

    # Get drive information
    $drive = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq "$DriveLetter\" }

    if ($drive) {
        return ($drive.Used + $drive.Free) -gt $SizeLimit
    } else {
        return $false
    }
}

# Function to check if the script is running with administrative privileges
function Test-IsAdmin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-USB {

        # Get all USB storage devices
        $usbDrives = Get-Disk | Where-Object { $_.BusType -eq 'USB' -and $_.Size -gt 0 }

        # Check if any USB drives were found
        if ($usbDrives.Count -eq 0) {
            Write-Host "No USB storage devices with media found."
            exit
        }
    
        Write-Host "USB storage devices with media found:"
        # List each USB drive with its partitions and drive letters
        foreach ($drive in $usbDrives) {
            $partitions = Get-Partition -DiskNumber $drive.Number -ErrorAction SilentlyContinue
            foreach ($partition in $partitions) {
                $logicalDisk = Get-Volume -Partition $partition
                $partitionSizeGB = [math]::Round($partition.Size / 1GB, 2)
                $driveLetter = if ($logicalDisk.DriveLetter) { $logicalDisk.DriveLetter } else { "No drive letter" }
                Write-Host "Disk $($drive.Number): $($drive.FriendlyName) - Partition $($partition.PartitionNumber) Size: ${partitionSizeGB}GB - Drive Letter: " -NoNewline
                Write-Host $driveLetter -ForegroundColor Green -NoNewline
                Write-Host ""
            }
        }

    }

function Get-DriveLetter {
    
    # Define the variable
    $FoundLetter = $false
    $sizeLimit = 64GB

    while ($true) {
        Write-Host -NoNewLine "Please enter a Drive Letter or press 'Escape' to exit: "
        
        # Detect key press
        $key = [System.Console]::ReadKey($true)
        
        # Check if 'Escape' key is pressed
        if ($key.Key -eq [ConsoleKey]::Escape) {
            Write-Host "`nExiting..."
            exit
        }

        # Read user input
        $driveLetter = $key.KeyChar.ToString()
        
        # Use default if no input is provided
        if ([string]::IsNullOrWhiteSpace($driveLetter)) {
            $driveLetter = $defaultDriveLetter
        }

        # Ensure the drive letter is uppercase and add the colon
        $driveLetter = ($driveLetter.ToUpper() + ":")

        # Check if the drive letter exists
        if (!(Test-Path "$driveLetter\")) {
            Write-Host
            Write-Host "The drive letter $driveLetter does not exist."
            continue
        }

        # Check if the drive has a Windows directory
        if (Has-WindowsDirectory -DriveLetter $driveLetter) {
            Write-Host
            Write-Host "The drive letter $driveLetter has a Windows directory and is likely a system drive."
            continue
        }

        # Check if the drive exceeds the size limit
        if (Is-DriveSizeExceeds -DriveLetter $driveLetter -SizeLimit $sizeLimit) {
            Write-Host
            Write-Host "The drive letter $driveLetter has more than 64 GB."
            continue
        }

        # If all checks pass, exit the loop
        Write-Host
        Write-Host "The drive letter $driveLetter is valid, not a system drive, and does not exceed 64 GB."
        $FoundLetter=$True
        break
    }

    # Get the Win32_DiskDrive associated with the drive letter
    $disk = Get-WmiObject Win32_DiskDrive -Filter "InterfaceType='USB'" | ForEach-Object {
        $drive = $_
        $partitions = Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($drive.DeviceID)'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        
        foreach ($partition in $partitions) {
            $logicalDisks = Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass = Win32_LogicalDiskToPartition"
            foreach ($logicalDisk in $logicalDisks) {
                if ($logicalDisk.DeviceID -eq $driveLetter) {
                    return $drive
                }
            }
        }
    }
    return $FoundLetter,$disk,$driveLetter
}


# Function to Format the drive
function Format-Drive {

    Get-USB

    $Results = Get-DriveLetter
    $FoundLetter = $Results[0]
    $disk = $Results[1]
    $driveLetter = $Results[2]

    if ($disk) {
        $diskNumber = $disk.Index
        Write-Host "The disk number for drive letter $driveLetter is: $diskNumber"
        # Now you can use $diskNumber in Diskpart
    } else {
        Write-Host "Disk with the drive letter $driveLetter not found or it's not a USB drive"
    }

    # Check if FoundLetter is set to True
    if ($FoundLetter -eq $true) {
        Write-Host "Are you sure you want to clean Disk $diskNumber? This will erase all data. Type 'Y' to confirm : " -NoNewline -ForegroundColor Yellow
        $confirmation = Read-Host 
        if ($confirmation -eq 'Y') {

$scriptContent = @"
select disk $diskNumber
list disk
attributes disk clear readonly
clean
convert MBR
create partition primary size=2048
active
format fs=FAT32 quick label="WinPE"
assign letter=P
create partition primary
format fs=NTFS quick label="Images"
assign letter=I
list partition
exit
"@
            
            $scriptContent | Out-File -FilePath "PrepareUSB.txt" -Encoding ASCII

            Write-Host "Running diskpart with PrepareUSB.txt..."

            # Run diskpart with param.txt
            Start-Process -FilePath "diskpart.exe" -ArgumentList "/s PrepareUSB.txt" -NoNewWindow -Wait

            # Remove-Item -Path "PrepareUSB.txt" -Force

            Write-Host "diskpart has completed."
        }
    } else {
        Write-Host "FoundLetter is not set to True. Exiting..."
    }

    return $driveLetter

}

function Select-ItemFromList {
    param (
        [Parameter(Mandatory=$true)]
        [object[]]$ItemList
    )
    
    # Display the items with an order number
    for ($i = 0; $i -lt $ItemList.Length; $i++) {
        Write-Host "$($i + 1): $($ItemList[$i])"
    }

    # Ask the user to choose a number
    $selectedNumber = 0
    do {
        $choice = Read-Host "Please choose a number between 1 and $($ItemList.Length)"
        $selectedNumber = $choice -as [int]
    }
    while ($selectedNumber -lt 1 -or $selectedNumber -gt $ItemList.Length)

    # Return the item corresponding to the number chosen
    return $ItemList[$selectedNumber - 1]
}

# Function
function Apply-PE {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WimFile,
        
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    
    # Check if the WIM file exists
    if (-not (Test-Path -Path $WimFile)) {
        Write-Host "The specified WIM file does not exist."
        return
    }
    
    $DestinationRoot = "${DriveLetter}\"
    
    # Check if the drive letter is valid
    if (-not ([System.IO.DriveInfo]::GetDrives().Name -contains $DestinationRoot)) {
        Write-Host "The specified drive letter is not valid."
        return
    }

    # Apply the image from the WIM file to the specified drive letter
    try {
        Write-Host "Applying image from $WimFile to drive $DriveLetter..."
        Dism /Apply-Image /ImageFile:$WimFile /Index:1 /ApplyDir:$DestinationRoot /Verify /NoRpFix
        Write-Host "Image applied successfully."
    }
    catch {
        Write-Host "An error occurred while applying the image:"
        Write-Host $_.Exception.Message
    }
}

# Function
function Get-ScriptFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourceDirectory,
        
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    
    # Ensure the source directory exists
    if (-not (Test-Path -Path $SourceDirectory)) {
        Write-Error "The specified source directory does not exist."
        return
    }
    
    # Ensure the drive letter ends with a colon and backslash
    $DestinationRoot = "${DriveLetter}\"

    # Ensure the drive letter is valid
    if (-not (Test-Path -Path $DestinationRoot)) {
        Write-Error "The specified drive letter is not valid."
        return
    }
    
    # Copy all files from the source directory to the root of the specified drive
    try {
        Get-ChildItem -Path $SourceDirectory -File |
        Copy-Item -Destination $DestinationRoot -Force
        Write-Host "Files copied successfully to $DestinationRoot"
    }
    catch {
        Write-Error "An error occurred while copying the files: $_"
    }
}

function Update-VariableInFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$VariableName,

        [Parameter(Mandatory = $true)]
        [string]$NewValue
    )

    # Check if the file exists
    if (-not (Test-Path -Path $FilePath)) {
        Write-Error "The specified file does not exist."
        return
    }

    # Read the file content
    $fileContent = Get-Content -Path $FilePath

    # Update the variable in the file content
    $updatedContent = $fileContent | ForEach-Object {
        if ($_ -match "^\s*$VariableName\s*=") {
            "$VariableName=$NewValue"
        } else {
            $_
        }
    }

    # Write the updated content back to the file
    Set-Content -Path $FilePath -Value $updatedContent
}

# Function
function Get-ModelWims {

    Write-Host "Copy the model WIMs to the key"
    $subdirectories = Get-ChildItem -Path $DEV_ROOT_WIM -Directory | Where-Object { $_.Name.Length -eq 7 }
    $subdirectoryNames = $subdirectories | Select-Object -ExpandProperty Name
    if ($subdirectoryNames) {
        $selectedModel = Select-ItemFromList -ItemList $subdirectoryNames    
    } else {
        write-host "No model directory found - Do the copy manually"
        return
    }

    Write-Host "Model $selectedModel chosen"

    $sourceDirectory = $DEV_ROOT_WIM + "\" + $selectedModel
    $targetDirectory = $NTFSLetter + "\"

    $inifile = $targetDirectory + "config.ini"
    Update-VariableInFile -FilePath $inifile -VariableName "ModelString" -NewValue $selectedModel

    # Get all .wim files in the source directory
    $wimFiles = Get-ChildItem -Path $sourceDirectory -Filter "*.wim"

    $totalFiles = $wimFiles.Count
    $fileCounter = 0

    foreach ($file in $wimFiles) {
        $fileCounter++
        $status = "Copying $($file.Name) to $targetDirectory"
        $percentComplete = ($fileCounter / $totalFiles) * 100
    
        # Show progress bar
        Write-Progress -Activity "Copying Files" -Status $status -PercentComplete $percentComplete
    
        Copy-Item -Path $file.FullName -Destination $targetDirectory -Force
    }
    # Hide the progress bar when done
    Write-Progress -Activity "Copying Files" -Completed

}

# Check for administrative privileges
if (-not (Test-IsAdmin)) {
    Write-Host "This script requires administrative privileges. Please run it as an administrator." -ForegroundColor Red
    exit
}

# Load external variable
$configData = Import-PowerShellDataFile -Path ".\config.psd1"
$PE_WIM = $configData["WinPE_WIM_Location"]
$DEV_ROOT_WIM = $configData["Device_Root_WIM_Location"]
$DRIVE_PE = ""

# Check validity of paths
if ( -not (Test-Path -Path $PE_WIM -PathType Leaf)) {
    Write-Host "The file path for WinPE WIM is not valid in config.psd1"
    exit
}

# Check if the directory path is valid and the directory exists
if ( -not (Test-Path -Path $DEV_ROOT_WIM -PathType Container)) {
    Write-Host "The directory path for the root of devices WIMs is not valid or the directory does not exist."
    exit
}

Write-Host "Do you want to format a new drive ? Y/N : " -NoNewline -ForegroundColor Yellow
$Response = Read-Host
# Check for administrative privileges
if ( $Response -eq "Y") {
    $DRIVE_PE = Format-Drive
    Apply-PE -WimFile $PE_WIM -DriveLetter $DRIVE_PE
} else {

    Write-Host "Do you want to Apply the PE content ? Y/N : " -NoNewline -ForegroundColor Yellow
    $Response = Read-Host
    if ( $Response -eq "Y") {

        Get-USB
        $FoundLetter = $False
        while (-not $FoundLetter) {
            $Results = Get-DriveLetter
            $FoundLetter = $Results[0]
            if ($FoundLetter)
            {
                $DRIVE_PE = $Results[2]
                Apply-PE -WimFile $PE_WIM -DriveLetter $DRIVE_PE
            }
        }
    }        
}


Write-Host "Please select the destination drive for scripts" -ForegroundColor Yellow

Get-USB
$Results = Get-DriveLetter
$FoundLetter = $Results[0]
$disk = $Results[1]
$NTFSLetter = $Results[2]

Get-ScriptFiles -SourceDirectory ".\Scripts" -DriveLetter $NTFSLetter

Write-Host "Do you want also to provide a set of model WIMs on drive $NTFSLetter ? Y/N : " -NoNewline -ForegroundColor Yellow
$Response = Read-Host
# Check for administrative privileges
if ( $Response -eq "Y") {
    Get-ModelWims
}

Write-Host "Merci d'avoir utilisé ProduceKey ... Fin des programmes" -ForegroundColor Yellow
