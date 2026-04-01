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

1. **Download Windows ISO** from Microsoft, then run `ExtractWim.ps1` to extract the Pro edition into `Windows_WIM_Root` automatically — or extract `install.wim` manually from the `sources` folder of the ISO
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
| `EditionIndex` | WIM edition index to apply (default `1`). Auto-set by `ProduceKey.ps1` when copying a multi-edition WIM. Use `dism /Get-ImageInfo /ImageFile:<wim>` to list available indices. |

> [!WARNING]
> If the ModelString doesn't match the device's SystemProductName, deployment will abort. This safety feature prevents accidental data loss. Always verify `DiskNumber` is correct - if the USB is disk 0, set `DiskNumber` to the internal disk number (often 1).

## config.psd1 (workstation - build-time paths)

Edit before running any workstation script:

```powershell
@{
    "WinPE_WIM_Location"       = "C:\Path\to\WinPE.wim"
    "Device_Root_WIM_Location" = "C:\Path\to\ModelWIMs"
    "Windows_WIM_Root"         = "C:\Path\to\WindowsWIMs"
}
```

| Key | Description |
|-----|-------------|
| `WinPE_WIM_Location` | Root directory of the WinPE library populated by `CreatePE.ps1` (one subdirectory per version/custom build, each containing a `boot.wim`) |
| `Device_Root_WIM_Location` | Root of the model WIM library (one subdirectory per model) |
| `Windows_WIM_Root` | Root of the vanilla Windows WIM library populated by `ExtractWim.ps1` (one subdirectory per Windows version) |

# USB Key Generation

## Prerequisites

- USB key 16GB+ (32GB recommended for multiple models)
- WinPE WIM file with custom startnet.cmd (see WinPE section)
- Windows ADK installed
- Administrator privileges

## Running ProduceKey

```powershell
# Interactive (no log file)
.\ProduceKey.ps1

# With log file
.\ProduceKey.ps1 -Log
```

The script guides you through five steps in order. Each step is optional — you can skip steps you don't need:

| Step | Action | Notes |
|------|--------|-------|
| 1 | **Format USB** | Creates dual-partition structure (FAT32 WinPE + NTFS Images) and optionally applies WinPE |
| 2 | **Apply WinPE** | Applies a PE from the library to an existing partition (if Step 1 was skipped) |
| 3 | **Update startnet.cmd** | Refreshes only `startnet.cmd` inside `boot.wim` on an existing WinPE partition — without re-applying the whole image |
| 4 | **Copy scripts** | Copies deployment scripts (`ApplyImage.bat`, `CaptureImage.bat`, `config.ini`, etc.) to the Images partition |
| 5 | **Copy model WIMs** | Copies WIM files for a selected model; auto-detects Pro edition index and updates `config.ini` |

A safety marker file (`DEPLOYKEY.marker`) is created to prevent accidental USB formatting.

> [!NOTE]
> PE driver injection (storage/NVMe drivers) is managed separately by `CreatePE.ps1` via the `PEDrivers\` folder. Inject drivers into the WinPE image before adding it to the library, not after applying it to a USB key.

Log file: `ProduceKey_YYYYMMDD_HHMMSS.log`

# Workstation Tooling Scripts

These scripts run on your workstation (not in WinPE) to build and maintain the WIM library used during key production. They all share a common module (`DeviceRebuild.psm1`) and read paths from `config.psd1`.

## GetDriver.ps1

Automated downloader for **SCCM Driver Packages** from the ASUS Support site. It resolves the model name, fetches the driver list via the ASUS API, and handles download and extraction.

```powershell
# Interactive
.\GetDriver.ps1 -Model "BM1403CDA"

# Automated download and expand
.\GetDriver.ps1 -Model "BM1403CDA" -OSVersion "WIN11_25H2" -AutoExpand $true

# Get latest version info as JSON
.\GetDriver.ps1 -Model "BM1403CDA" -GetLatestVersion
```

| Parameter | Description |
|-----------|-------------|
| `-Model` | Device model name (e.g., `BM1403` or `BM1403CDA`) |
| `-OSVersion` | Target OS version string (e.g., `WIN11_25H2`). Defaults to `25H2` |
| `-Version` | Specific version string (e.g., `01.00`) to download |
| `-Destination` | Directory to save the download. Defaults to current directory |
| `-AutoExpand` | If set (`$true` or `$false`), runs non-interactively. If `$true`, extracts and deletes the ZIP |
| `-GetLatestVersion` | Returns a JSON object with `{Version, OSVersion}` and exits |

**Automated Workflow Example:**
```powershell
# 1. Get latest info
$info = .\GetDriver.ps1 -Model BM1403CDA -GetLatestVersion | ConvertFrom-Json

# 2. Download and expand using that exact info
.\GetDriver.ps1 -Model BM1403CDA -Version $info.Version -OSVersion $info.OSVersion -AutoExpand $true
```

## ExtractWim.ps1

Extracts the **Pro** and **Pro Education** editions from a Windows ISO and saves them as a single version-named WIM file under `Windows_WIM_Root`.

```powershell
# With ISO path as parameter
.\ExtractWim.ps1 -IsoPath "C:\ISOs\Win11_24H2.iso"

# Without parameter — opens a file picker dialog
.\ExtractWim.ps1

# With log file
.\ExtractWim.ps1 -IsoPath "C:\ISOs\Win11_24H2.iso" -Log
```

**Output:** `<Windows_WIM_Root>\<Version>\<Version>.wim`

```
C:\Scratch\WindowsWIMs\
├── W11-24H2\
│   └── W11-24H2.wim     (contains: Windows 11 Pro + Windows 11 Pro Education)
├── W11-25H2\
│   └── W11-25H2.wim
└── W10-22H2\
    └── W10-22H2.wim
```

The version name (`W11-24H2`, `W10-22H2`, etc.) is automatically detected from the ISO's build number. If the build is unknown, the script prompts you to enter the name manually.

**Supported version detection:**

| Build | Version name |
|-------|-------------|
| 19041–19045 | W10-2004 … W10-22H2 |
| 22000 | W11-21H2 |
| 22621 | W11-22H2 |
| 22631 | W11-23H2 |
| 26100 | W11-24H2 |
| 26200 | W11-25H2 |
| other | Prompts for name |

> [!NOTE]
> Both `install.wim` and `install.esd` ISOs are supported. `.esd` export is slower due to the encrypted compression format.

> [!TIP]
> The resulting WIM can be renamed to `ModelName-OS.wim` and placed in your `Device_Root_WIM_Location` model directory, then used by `ProduceKey.ps1` to populate a USB key. `ProduceKey.ps1` will automatically detect the Pro edition index during the WIM copy step.

# WIM File Naming Convention

**Naming format:** `ModelName-PartitionType.wim`

| Partition Type | Description |
|---------------|-------------|
| OS | Windows system partition |
| RECOVERY | Windows Recovery Environment |
| MYASUS | ASUS vendor partition (optional) |
| SYSTEM | EFI system partition (rebuilt by bcdboot) |

## Workstation Source Directory Structure

Place model directories under `Device_Root_WIM_Location`. The directory name is used as the model string unless overridden by a `model.ini` file inside it.

**Factory capture layout** (OOBE capture or dirty snapshot):
```
Device_Root_WIM_Location/
├── B9450FA/
│   ├── B9450FA-OS.wim        (or .swm for split images)
│   ├── B9450FA-OS2.swm       (split WIM parts, if any)
│   ├── B9450FA-RECOVERY.wim
│   └── B9450FA-MYASUS.wim
└── B5402CB/
    └── ...
```

**Vanilla OS + driver injection layout** (Use Case 1):
```
Device_Root_WIM_Location/
└── B9450FA/
    ├── B9450FA-OS.wim        (renamed from Microsoft install.wim)
    └── Drivers/              (optional - triggers driver injection)
        ├── Chipset/
        │   └── *.inf
        ├── Graphics/
        │   └── *.inf
        ├── Network/
        │   └── *.inf
        ├── Audio/
        │   └── *.inf
        └── Rapid Storage/    (NVMe/RAID devices - also injected into winre.wim)
            └── *.inf
```

> [!NOTE]
> If the directory name does not match the device's `ModelString`, create a `model.ini` file inside the directory to override it:
> ```ini
> ModelString=B9450FA
> ```
> This lets you use a descriptive folder name (e.g., `ExpertBook B9`) while keeping the correct model string for deployment.

## LNK Redirect System

To avoid duplicating large WIM files across model directories, you can place marker files that redirect `ProduceKey.ps1` to a shared source. All files are renamed on the fly with the target model prefix when copied to the USB key.

### Level 1 — Share WIMs between models

Create an empty file named `LNK-<SourceModel>.txt` in the model directory. `ProduceKey.ps1` will pull all WIM files from the `<SourceModel>` directory, rename them with the target model prefix, and copy them to the USB key. The target model directory typically holds:

- `LNK-<SourceModel>.txt` — redirect marker
- `Drivers\` — model-specific drivers (always from here, never from the source)
- `ModelA-SYSTEM.wim`, `ModelA-RECOVERY.wim`, `ModelA-MYASUS.wim` — partition WIMs that differ from the source (optional, copied as-is)

```
Device_Root_WIM_Location/
├── B9450FA/
│   ├── B9450FA-OS.wim         (self-contained with all WIMs)
│   ├── B9450FA-RECOVERY.wim
│   └── B9450FA-MYASUS.wim
└── B5402CB/
    ├── LNK-B9450FA.txt        (OS comes from B9450FA\)
    ├── B5402CB-RECOVERY.wim   (model-specific partitions kept here)
    ├── B5402CB-MYASUS.wim
    └── Drivers\               (model-specific drivers)
```

When B5402CB is selected: `B9450FA-OS.wim` is copied as `B5402CB-OS.wim`.

### Level 2 — Use vanilla Windows WIM as OS

Create an empty file named `LNK-OS-<Version>.txt` inside the Level 1 source directory (e.g., `B9450FA\`). This further redirects the OS WIM to `Windows_WIM_Root\<Version>\<Version>.wim` (the output of `ExtractWim.ps1`).

```
Device_Root_WIM_Location/
├── B9450FA/
│   ├── LNK-OS-W11-25H2.txt   (OS from Windows_WIM_Root\W11-25H2\W11-25H2.wim)
│   ├── B9450FA-RECOVERY.wim
│   └── B9450FA-MYASUS.wim
└── B5402CB/
    ├── LNK-B9450FA.txt        (inherit from B9450FA, which itself uses W11-25H2)
    ├── B5402CB-RECOVERY.wim
    └── Drivers\

Windows_WIM_Root/
└── W11-25H2/
    └── W11-25H2.wim           (Pro + Pro Education, output of ExtractWim.ps1)
```

When either model is selected, `W11-25H2.wim` is copied to the USB key as `<ModelString>-OS.wim`, and `ProduceKey.ps1` auto-detects the Pro edition index.

| Scenario | OS source | Other WIMs |
|----------|-----------|------------|
| No LNK | Model dir | Model dir |
| `LNK-<Src>.txt` only | `<Src>\` dir (renamed) | Model dir + `<Src>\` dir |
| `LNK-OS-<Ver>.txt` only (in model dir) | `Windows_WIM_Root\<Ver>\` | Model dir |
| `LNK-<Src>.txt` + `LNK-OS-<Ver>.txt` (in `<Src>\`) | `Windows_WIM_Root\<Ver>\` | Model dir + `<Src>\` dir (minus OS) |

`ProduceKey.ps1` copies WIM files to the root of `I:\` and the `Drivers\` subfolder to `I:\Drivers\<ModelString>\` automatically.

## USB Key NTFS Partition Structure (I:\\)

This is the layout on the USB storage partition (`I:\`) as seen by `ApplyImage.bat` at boot time. All files must be at the root — this is where `startnet.cmd` changes to before calling the deployment script.

**Vanilla OS + driver injection:**
```
I:\                            (USB NTFS storage partition)
├── config.ini
├── DEPLOYKEY.marker
├── ApplyImage.bat
├── CaptureImage.bat
├── B9450FA-OS.wim             (or B9450FA-OS.swm + B9450FA-OS2.swm for split)
└── Drivers\
    └── B9450FA\               (must match ModelString in config.ini)
        ├── Chipset\
        │   └── *.inf
        ├── Graphics\
        │   └── *.inf
        ├── Network\
        │   └── *.inf
        ├── Audio\
        │   └── *.inf
        └── Rapid Storage\
            └── *.inf
```

**Factory image deployment:**
```
I:\
├── config.ini
├── DEPLOYKEY.marker
├── ApplyImage.bat
├── CaptureImage.bat
├── B9450FA-OS.wim             (or split .swm files)
├── B9450FA-RECOVERY.wim       (optional)
└── B9450FA-MYASUS.wim         (optional)
```

> [!TIP]
> When manually preparing a USB key without using `ProduceKey.ps1`, copy the scripts from the `Scripts\` folder and the WIM files directly to `I:\` (root), and place drivers under `I:\Drivers\<ModelString>\`.

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

DeviceRebuild requires a WinPE (Windows Preinstallation Environment) WIM file to boot the USB key. 

- **`CreatePE.ps1`** — builds a fully customized WinPE using the Windows ADK. Supports PE driver injection, keyboard layout, PowerShell, and manual customization.

## Building a customized WinPE with CreatePE

`CreatePE.ps1` requires the Windows ADK + WinPE Add-on and operates in two independent phases:

```powershell
# Run as Administrator
.\CreatePE.ps1
```

**Phase A — Build PE WIM** (no USB required):
1. Select or create an output directory in the PE library
2. `copype` creates a fresh WinPE staging environment
3. Inject PE drivers from `PEDrivers\` (optional, multi-select)
4. Set keyboard layout, add PowerShell support, pause for manual edits
5. Capture result as `boot.wim` in the PE library

**Phase B — Apply PE to USB** (optional, runs after Phase A or standalone):
- Picks any `boot.wim` from the PE library
- Applies it to a USB WinPE partition using DISM
- Injects the latest `startnet.cmd` automatically

> [!NOTE]
> Phase A and Phase B are fully independent. You can build a PE without any USB connected, then apply it later. You can also apply an existing PE from the library without rebuilding.

## Injecting drivers into WinPE (PEDrivers)

To inject storage, NVMe, or other drivers into WinPE at build time, create a `PEDrivers\` folder next to the PE library root (`WinPE_WIM_Location`):

```
<WinPE_WIM_Location>\
├── PEDrivers\               <- driver packages for WinPE injection
│   ├── IntelRST\            <- one subfolder per driver package
│   │   └── *.inf
│   ├── Chipset\
│   │   └── *.inf
│   └── NetworkDriver\
│       └── *.inf
├── W11-24H2\
│   └── boot.wim
└── CustomPE\
    └── boot.wim
```

When `CreatePE.ps1` runs Phase A, it detects all subfolders in `PEDrivers\` and presents a toggleable checklist. Type a number to toggle a package on/off, `A` to select all, then press Enter to confirm.

> [!TIP]
> For NVMe/RAID devices (e.g., Intel RST), inject the storage driver here so WinPE can detect the internal disk. This replaces the old per-model driver scanning that was previously done at USB production time.

## Updating startnet.cmd on an existing USB

**Via ProduceKey.ps1** (Step 3 — recommended):
```powershell
.\ProduceKey.ps1
# Choose "N" to skip format, "N" to skip WinPE apply, "Y" to update startnet.cmd
# Select the WinPE partition (e.g. P:)
```

**Via batch script** (manual):
```shell
Scripts\UpdateStartnet.bat P
```

Where `P` is the WinPE partition drive letter.

**Manual DISM** (advanced):
```shell
Dism /Mount-Image /ImageFile:"P:\sources\boot.wim" /index:1 /MountDir:"C:\WinPE_mount"
copy Scripts\startnet.cmd "C:\WinPE_mount\Windows\System32\startnet.cmd"
Dism /Unmount-Image /MountDir:"C:\WinPE_mount" /commit
```

# Troubleshooting

## USB Key Lost Partitions
If DiskNumber was incorrectly set to the USB disk number, the USB would be wiped. Always verify with `diskpart` → `list disk` before deployment.

## Model Mismatch Error
Check that `ModelString` in config.ini is contained within the device's SystemProductName (visible in log file).

## WinRE Not Working
Check ApplyImage log for ReAgentC errors. Ensure RECOVERY.wim contains `Recovery\WindowsRE\winre.wim`.

## Drivers Not Injecting
Verify folder structure: `Drivers\ModelString\` with `.inf` files. Check log for DISM errors.
