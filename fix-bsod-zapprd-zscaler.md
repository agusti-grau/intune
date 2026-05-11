# Fixing Windows BSOD Loop: `DRIVER_UNLOADED_WITHOUT_CANCELLING_PENDING_OPERATIONS` (0xCE) caused by `zapprd.sys`

## Problem

After uninstalling **Zscaler Client Connector**, Windows enters a reboot loop with the following Blue Screen:

```
Your device ran into a problem and needs to restart.
We'll restart for you.

Stop code: DRIVER_UNLOADED_WITHOUT_CANCELLING_PENDING_OPERATIONS (0xCE)
What failed: zapprd.sys
```

## Cause

`zapprd.sys` is the **Zscaler App Profiler Driver**, a kernel-mode component installed by Zscaler Client Connector. When the application is uninstalled but the driver and its registry service entry remain, Windows still tries to load the driver at boot. Without its parent service intact, the driver unloads while it still has pending I/O operations, triggering bug check `0xCE` and a boot loop.

## Prerequisites

- Ability to boot into **Safe Mode** or reach the **Windows Recovery Environment (WinRE)**.
- Administrator-level access — if not available within Windows, **WinRE** runs as `SYSTEM` and bypasses UAC, which is the fallback path used in this guide.

---

## Step 1 — Try Safe Mode first (if you have admin rights)

Boot into Safe Mode and open **Command Prompt** or **PowerShell as Administrator**.

### 1.1 Check what Zscaler components remain

```cmd
sc query ZAPPRD
sc query ZSATunnel
sc query ZSAService
sc query ZSATrayManager
dir C:\Windows\System32\drivers\zapprd.sys
```

### 1.2 Stop and delete leftover services

Ignore errors for any service that no longer exists.

```cmd
sc stop ZAPPRD
sc delete ZAPPRD
sc stop ZSATunnel
sc delete ZSATunnel
sc stop ZSAService
sc delete ZSAService
sc stop ZSATrayManager
sc delete ZSATrayManager
```

### 1.3 Remove the driver package from the DriverStore

```cmd
pnputil /enum-drivers | findstr /i zapp
```

The output shows an `oemNN.inf` entry. Remove it:

```cmd
pnputil /delete-driver oemNN.inf /uninstall /force
```

### 1.4 Delete the driver file

```cmd
del /f C:\Windows\System32\drivers\zapprd.sys
```

If you get **Access Denied** or you cannot elevate Command Prompt, skip to **Step 2 (WinRE)**.

---

## Step 2 — Fix from WinRE (recommended when permissions fail)

The WinRE Command Prompt runs as `SYSTEM`, so it bypasses user permissions, UAC, and file locks.

### 2.1 Boot into WinRE

From Safe Mode: **Settings → System → Recovery → Advanced startup → Restart now**.
Or force three failed boots (cut power during boot three times) — Windows will open WinRE automatically.

Then navigate: **Troubleshoot → Advanced options → Command Prompt**.

### 2.2 Identify the Windows drive letter

In WinRE, the Windows volume is often **not** `C:`. Identify it:

```cmd
diskpart
list volume
exit
```

Pick the volume labeled **Windows** (usually the largest NTFS volume). The examples below assume it is `D:` — replace as needed.

### 2.3 Delete the driver file

```cmd
del /f /a D:\Windows\System32\drivers\zapprd.sys
```

If access is denied, take ownership first:

```cmd
takeown /f D:\Windows\System32\drivers\zapprd.sys
icacls D:\Windows\System32\drivers\zapprd.sys /grant Administrators:F
del /f D:\Windows\System32\drivers\zapprd.sys
```

### 2.4 Disable the driver service in the offline registry

This is the critical step that stops Windows from trying to load the driver on the next boot.

```cmd
reg load HKLM\OFFSYS D:\Windows\System32\config\SYSTEM
reg delete "HKLM\OFFSYS\ControlSet001\Services\ZAPPRD" /f
reg delete "HKLM\OFFSYS\ControlSet002\Services\ZAPPRD" /f
reg unload HKLM\OFFSYS
```

Ignore errors for any `ControlSet00X` that doesn't exist on your system.

### 2.5 Reboot

```cmd
exit
```

Then choose **Continue** to boot into Windows normally.

---

## Step 3 — Post-fix cleanup (back in normal Windows)

Open an **elevated** Command Prompt or PowerShell and run:

```cmd
pnputil /enum-drivers | findstr /i zapp
pnputil /delete-driver oemNN.inf /uninstall /force
```

```cmd
rmdir /s /q "C:\Program Files\Zscaler"
rmdir /s /q "C:\ProgramData\Zscaler"
reg delete "HKLM\SOFTWARE\Zscaler Inc." /f
reg delete "HKLM\SYSTEM\CurrentControlSet\Services\ZAPPRD" /f
```

---

## Verification

The fix is successful when:

- The system boots into Windows without a BSOD.
- `dir C:\Windows\System32\drivers\zapprd.sys` returns *File Not Found*.
- `sc query ZAPPRD` returns *The specified service does not exist*.
- `pnputil /enum-drivers | findstr /i zapp` returns no matches.

---

## Notes

- The same approach works for other orphaned third-party kernel drivers that survive a botched uninstall (antivirus, VPNs, endpoint agents). Adjust the driver name and service name accordingly.
- The proper long-term fix is to use the vendor's official uninstaller (Zscaler ships a removal utility for admins). The procedure above is the recovery path when uninstallation has already left the system unbootable.
- Always create a system restore point or full backup before editing the offline registry.

---

## References

- Microsoft — Bug Check 0xCE: `DRIVER_UNLOADED_WITHOUT_CANCELLING_PENDING_OPERATIONS`
- Microsoft — `pnputil` command-line reference
- Microsoft — `reg load` / `reg unload` for offline registry editing
