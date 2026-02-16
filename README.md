# DeviceRebuild

# Purpose

Automatically rebuild a Windows Device based on a OOBE WIM capture of the model.

# Features

- **Dual-partition USB keys**: WinPE boot partition (FAT32) + Images partition (NTFS)
- **Model validation**: Prevents deployment to wrong device models
- **Split WIM support**: Handles large OS images split into .swm files
- **Driver injection**: Optional offline driver injection for vanilla Windows images
- **OOBE BypassNRO**: Skip the network requirement during Windows setup
- **Comprehensive logging**: Timestamped log files for troubleshooting
- **Safety checks**: Prevents accidental formatting of the USB key itself (DEPLOYKEY.marker)
- **WinRE configuration**: Automatic Windows Recovery Environment setup with Rapid Storage driver injection
- **Push Button Reset**: Automatic configuration of ResetConfig.xml for factory reset capability

# Use Cases

## 1. Clean Install from Microsoft ISO (with Driver Injection)

Rebuild a device from a vanilla Windows image downloaded from Microsoft, with manufacturer drivers injected offline.

**When to use:** You want a clean Windows installation with the correct drivers for a specific hardware model.

**Steps:**

1. **Download Windows ISO** from Microsoft and extract the `install.wim` from the `sources` folder
2. **Rename the WIM** to match the naming convention: `ModelName-OS.wim` (e.g., `B9450FA-OS.wim`)
3. **Prepare drivers** from the manufacturer website:
   - Download driver packages for your model (Chipset, Graphics, Network, Audio, Storage, etc.)
   - Extract them so that `.inf` files are accessible
   - Place them in a folder structure on the USB key:
     ```
     I:\Drivers\B9450FA\
     ├── Chipset\
     │   └── *.inf
     ├── Graphics\
     │   └── *.inf
     ├── Network\
     │   └── *.inf
     ├── Audio\
     │   └── *.inf
     └── Rapid Storage\    (important for NVMe/RAID devices)
         └── *.inf
     ```
4. **Place the WIM file** at the root of the USB NTFS partition (`I:\`)
5. **Edit `config.ini`**:
   ```ini
   ModelString=B9450FA
   targetScript=ApplyImage.bat
   DiskNumber=0
   BypassNRO=1
   ```
6. **Boot the target device** from the USB key and confirm deployment
7. The script will: partition the disk, apply the OS image, inject drivers, configure boot files, set up WinRE, and shut down
8. On next boot the device enters OOBE for a fresh Windows setup

> [!NOTE]
> For NVMe/RAID devices, include Rapid Storage drivers in `Drivers\ModelName\Rapid Storage\`. These will be automatically injected into `winre.wim` so that Windows Recovery can access the disk.

> [!NOTE]
> Since there is no captured RECOVERY or MYASUS partition, the script will skip those and set up WinRE from the `winre.wim` included in the OS image.

## 2. Factory Capture (Before OOBE)

Capture a brand-new device at first boot, before starting OOBE. This preserves the exact factory state including all manufacturer partitions (Recovery, vendor partitions like MyASUS). You can later re-apply these images to restore the device to its factory condition.

**When to use:** You just received a new device and want to create a factory backup before using it, so you can reset to factory state at any time without relying on the manufacturer's recovery solution.

**Steps:**

1. **Do NOT start the device normally** - interrupt the boot before OOBE begins
2. **Boot from the USB key** (e.g., _Reboot + Esc_ or _F2 + F8_ on ASUS devices)
3. **Identify the disk layout** using `diskpart` → `list disk` → `list volume` to find the correct drive letters for each partition
4. **Create `Mount-AllLetters.txt`** (diskpart script) to assign the expected drive letters: W: (Windows), S: (System), R: (Recovery), M: (MyASUS)
5. **Edit `config.ini`**:
   ```ini
   ModelString=B9450FA
   targetScript=CaptureImage.bat
   DiskNumber=0
   ```
6. **Boot from USB and confirm** - the script captures all partitions to individual WIM files:
   - `B9450FA-OS.wim` - Windows partition
   - `B9450FA-SYSTEM.wim` - EFI System partition
   - `B9450FA-RECOVERY.wim` - Recovery partition
   - `B9450FA-MYASUS.wim` - Vendor partition
7. **Copy the WIM files** to your storage repository for future use

> [!TIP]
> No driver injection is needed when re-applying these images since they already contain all manufacturer drivers from the factory installation.

> [!TIP]
> If the OS WIM is too large for FAT32 constraints (>4GB), split it for storage: `Dism /Split-Image /ImageFile:B9450FA-OS.wim /SWMFile:B9450FA-OS.swm /FileSize:4000`. ApplyImage.bat automatically detects and handles split images.

## 3. Dirty Snapshot (For Benchmarking)

Capture a device after installing benchmark tools and optimizing settings, without running sysprep. This creates a "dirty" snapshot that can be quickly re-applied to reset the device to a known benchmark-ready state between test runs.

**When to use:** You have a device configured specifically for benchmarking (tools installed, settings optimized, background services disabled, etc.) and need to reset it to this exact state before each new test run.

**Steps:**

1. **Set up the device** for benchmarking:
   - Install all benchmark tools and utilities
   - Optimize Windows settings (disable updates, telemetry, background apps, etc.)
   - Configure power plans, display settings, etc.
   - Run through one test cycle to ensure everything works
2. **Shut down the device** cleanly
3. **Boot from the USB key**
4. **Edit `config.ini`** to capture:
   ```ini
   ModelString=B9450FA
   targetScript=CaptureImage.bat
   DiskNumber=0
   ```
5. **Run the capture** - all partitions are saved as WIM files
6. **After each benchmark run**, re-apply the snapshot:
   - Change `config.ini`: `targetScript=ApplyImage.bat`
   - Boot from USB - the device is restored to the exact benchmark-ready state
   - No OOBE, no driver installation, no reconfiguration needed

> [!NOTE]
> This is called a "dirty" snapshot because no sysprep is performed. The captured image retains the machine's SID, user profiles, and installed software exactly as configured. This is perfectly fine for benchmarking where you're always re-applying to the same physical device.

# Configuration

## config.ini (USB key - runtime settings)

Edit this file before each operation:

| Variable | Description |
|----------|-------------|
| `ModelString` | Device model name (typically 6-8 characters, e.g., `B9450FA`) |
| `targetScript` | Script to execute: `ApplyImage.bat` or `CaptureImage.bat` |
| `DiskNumber` | Target disk number (use `diskpart` → `list disk` to find it) |
| `BypassNRO` | Set to `1` to skip network requirement during OOBE (optional) |

> [!WARNING]
> If the ModelString doesn't match the device's SystemProductName, deployment will abort. This safety feature prevents accidental data loss. Always verify `DiskNumber` is correct - if the USB is disk 0, set `DiskNumber` to the internal disk number (often 1).

## config.psd1 (workstation - build-time paths)

Edit before running ProduceKey:

```powershell
@{
    "WinPE_WIM_Location" = "C:\Path\to\WinPE.wim"
    "Device_Root_WIM_Location" = "C:\Path\to\ModelWIMs"
}
```

# USB Key Generation

## Prerequisites

- USB key 16GB+ (32GB recommended for multiple models)
- WinPE WIM file with custom startnet.cmd (see WinPE section)
- Windows ADK installed
- Administrator privileges

## Running ProduceKey

```powershell
.\ProduceKey.ps1
```

The script guides you through:
1. **Format USB** (optional) - Creates dual-partition structure
2. **Apply WinPE** (optional) - Installs WinPE to boot partition
3. **Copy scripts** - Copies deployment scripts to USB
4. **Copy model WIMs** (optional) - Copies WIM files for a specific model

A safety marker file (`DEPLOYKEY.marker`) is created to prevent accidental USB formatting.

Log file: `ProduceKey_YYYYMMDD_HHMMSS.log`

# WIM File Naming Convention

Place WIM files in model-specific subdirectories under `Device_Root_WIM_Location`:

```
DeviceWIMs/
├── B9450FA/
│   ├── B9450FA-OS.wim        (or .swm for split images)
│   ├── B9450FA-OS2.swm       (split WIM parts)
│   ├── B9450FA-RECOVERY.wim
│   └── B9450FA-MYASUS.wim
└── B5402CB/
    └── ...
```

**Naming format:** `ModelName-PartitionType.wim`

| Partition Type | Description |
|---------------|-------------|
| OS | Windows system partition |
| RECOVERY | Windows Recovery Environment |
| MYASUS | ASUS vendor partition (optional) |
| SYSTEM | EFI system partition (rebuilt by bcdboot) |

## Split WIM Support (SWM)

For large OS images, use DISM to split:

```shell
Dism /Split-Image /ImageFile:B9450FA-OS.wim /SWMFile:B9450FA-OS.swm /FileSize:4000
```

This creates `B9450FA-OS.swm`, `B9450FA-OS2.swm`, etc. ApplyImage.bat automatically detects and uses split images.

# Partition Layout Created by ApplyImage

| # | Label | Size | Type | Purpose |
|---|-------|------|------|---------|
| 1 | System | 260MB | FAT32/EFI | UEFI boot files |
| 2 | (MSR) | 16MB | Reserved | Microsoft Reserved |
| 3 | Windows | * | NTFS | OS installation |
| 4 | Recovery | 2GB | NTFS | WinRE (winre.wim) |
| 5 | MyASUS | * | NTFS | Vendor partition |

# Log Files

All operations create timestamped log files on the USB key:

| Log File | Description |
|----------|-------------|
| `Startnet_*.log` | WinPE boot, config loading, model validation |
| `ApplyImage_*.log` | Partition creation, image deployment, driver injection |
| `CaptureImage_*.log` | Partition capture operations |

# Safety Features

## USB Key Protection

The script creates `DEPLOYKEY.marker` on the USB key. Before partitioning, ApplyImage.bat:
1. Checks all partitions on the target disk for this marker
2. If found, **aborts** with clear error message
3. Prevents accidental USB key destruction

## Model Validation

startnet.cmd validates that `ModelString` appears in the device's `SystemProductName` registry value before proceeding. Deployment is aborted on mismatch.

## DiskNumber Validation

`DiskNumber` must be explicitly set in config.ini. Empty or undefined values cause the script to abort.

## User Confirmation

Before executing the target script, startnet.cmd displays a summary and asks for explicit `Y/N` confirmation.

# WinPE WIM Customization

DeviceRebuild requires a WinPE (Windows Preinstallation Environment) WIM file to boot the USB key. This project cannot distribute the WIM file, but you can build one yourself using the Windows ADK.

## Building a WinPE WIM

1. **Install the Windows ADK** and the **WinPE add-on** from Microsoft
2. **Open the Deployment and Imaging Tools Environment** as Administrator
3. **Create a WinPE working directory:**
   ```shell
   copype amd64 C:\WinPE_amd64
   ```
4. The WinPE WIM is located at `C:\WinPE_amd64\media\sources\boot.wim`
5. **Set `WinPE_WIM_Location`** in `config.psd1` to point to this file (or copy it to a permanent location first)

For complete instructions, see Microsoft's official documentation:
[Create bootable WinPE media](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/winpe-create-usb-bootable-drive)

## Manual Updating startnet.cmd on Existing USB

To manually replace the startnet.cmd inside the WinPE WIM on an already-built USB key:

```shell
# Mount the WinPE WIM from the USB boot partition
Dism /Mount-Image /ImageFile:"P:\sources\boot.wim" /index:1 /MountDir:"C:\WinPE_mount"

# Copy the updated startnet.cmd from this repo
copy Scripts\startnet.cmd "C:\WinPE_mount\Windows\System32\startnet.cmd"

# Unmount and commit changes
Dism /Unmount-Image /MountDir:"C:\WinPE_mount" /commit
```

## Automatic Update startnet.cmd on Existing USB

Use the helper script to automate the steps above:

```shell
# Run as Administrator
Scripts\UpdateStartnet.bat P
```

Where `P` is the WinPE partition drive letter.

# Troubleshooting

## USB Key Lost Partitions
If DiskNumber was incorrectly set to the USB disk number, the USB would be wiped. Always verify with `diskpart` → `list disk` before deployment.

## Model Mismatch Error
Check that `ModelString` in config.ini is contained within the device's SystemProductName (visible in log file).

## WinRE Not Working
Check ApplyImage log for ReAgentC errors. Ensure RECOVERY.wim contains `Recovery\WindowsRE\winre.wim`.

## Drivers Not Injecting
Verify folder structure: `Drivers\ModelString\` with `.inf` files. Check log for DISM errors.
