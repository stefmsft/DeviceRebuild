# WinPE Tooling Setup — Windows ADK + WinPE Add-on

ExtractPE.ps1 requires the **Windows Assessment and Deployment Kit (ADK)** and its
**WinPE Add-on** to build bootable WinPE partition WIMs. This document explains
what to install, where to get it, and what to expect.

---

## Why ADK is required

Building a bootable WinPE partition is not simply exporting a WIM from an ISO.
A bootable WinPE FAT32 partition requires:

- `bootmgr` / `bootmgr.efi` — boot managers
- `Boot\BCD` — Boot Configuration Data **configured for standalone WinPE boot**
  (not Windows Setup — that is what an ISO BCD is configured for)
- `EFI\Boot\bootx64.efi` — UEFI boot entry point
- `EFI\Microsoft\Boot\BCD` — UEFI BCD
- `sources\boot.wim` — the WinPE OS image

The ADK provides the pre-built **media directory** (`amd64\Media\`) containing all
of the above with a correctly configured BCD. There is no reliable way to produce
that BCD from scratch without ADK tools.

---

## What to download

Install these two packages **in order**, on any Windows 10/11 machine:

### 1 — Windows ADK 10.1.26100.2454 (December 2024)
> Supports Windows 11 25H2, 24H2, and all earlier versions.

**Download:** https://go.microsoft.com/fwlink/?linkid=2289980

Filename: `adksetup.exe`

During installation, select **only**:
- ✅ **Deployment Tools**

Everything else is optional and not required for ExtractPE.ps1.

### 2 — Windows PE Add-on for ADK 10.1.26100.2454
Must be installed after the ADK above.

**Download:** https://go.microsoft.com/fwlink/?linkid=2289981

Filename: `adkwinpesetup.exe`

During installation, select:
- ✅ **Windows Preinstallation Environment (Windows PE)**

---

## What gets installed

Default install root:
```
C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\
```

Files used by ExtractPE.ps1:

| Path | Purpose |
|------|---------|
| `...\Windows Preinstallation Environment\amd64\Media\` | Pre-built boot structure (BCD, bootmgr, EFI) |
| `...\Windows Preinstallation Environment\amd64\en-us\winpe.wim` | Confirmation that WinPE add-on is installed |

ExtractPE.ps1 copies the `amd64\Media\` tree into a staging directory, then
injects the version-specific WinPE from your ISO as `sources\boot.wim`, and
captures the result as the final `boot.wim`.

---

## How ExtractPE.ps1 uses ADK (no manual steps required)

Once the ADK and WinPE Add-on are installed, run ExtractPE.ps1 normally:

```powershell
.\ExtractPE.ps1
```

It will:
1. Detect the ADK media directory automatically
2. Select a Windows ISO (file dialog or `-IsoPath` parameter)
3. Auto-detect the Windows version from the ISO build number
4. Build `<WinPE_WIM_Location>\<Version>\boot.wim`

---

## Manual reference — what copype.cmd does

The ADK ships `copype.cmd` (at `...\Windows Preinstallation Environment\copype.cmd`).
ExtractPE.ps1 replicates its behaviour directly without calling it, to avoid
dependency on the ADK environment shell. For reference:

```cmd
copype amd64 C:\WinPE_amd64
```

Produces:
```
C:\WinPE_amd64\
  media\                        ← the partition content
    bootmgr
    bootmgr.efi
    Boot\BCD                    ← WinPE-configured BCD (BIOS)
    Boot\boot.sdi
    EFI\Boot\bootx64.efi
    EFI\Microsoft\Boot\BCD      ← WinPE-configured BCD (UEFI)
    sources\boot.wim            ← ADK minimal WinPE (replaced by ExtractPE.ps1)
  mount\                        ← empty, for dism /Mount-Image use
```

ExtractPE.ps1 copies `amd64\Media\` directly instead of calling copype, which
achieves the same result without needing the ADK environment variables set.

---

## Verifying your installation

Open a regular PowerShell (Admin) and run:

```powershell
$adkRoot = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"
Test-Path "$adkRoot\Windows Preinstallation Environment\amd64\Media\Boot\BCD"
Test-Path "$adkRoot\Windows Preinstallation Environment\amd64\en-us\winpe.wim"
```

Both should return `True`. If either returns `False`, re-run the relevant installer.
