# Installation

Install **`func_SMILES2IMG`** across Microsoft Excel, Word, and PowerPoint in 10 seconds—**no administrator rights required**.

## 1. 10-Second Auto-Install (PowerShell)

1. Close any open Excel, Word, or PowerPoint windows.
2. Open **PowerShell** from the Windows Start menu, paste the command below, and press <kbd>Enter</kbd>:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip What's Included in `v1.9.0`
Running the command above automatically detects your Office bitness (32-bit vs. 64-bit), downloads **`func_SMILES2IMG v1.9.0`**, unblocks files (`Unblock-File`), and registers the add-in in HKCU for Excel, Word, and PowerPoint—featuring unified `Chemical Structure >` context menus, 49 Excel functions, Ketcher 2D, Mol\* 3D, and 11-column `.sdf` import.
:::

## 2. Restart Office & Verify the `[SMILES]` Ribbon Tab

1. **Completely close and reopen** Excel, Word, or PowerPoint (add-ins load on startup).
2. Confirm that the **`[SMILES]`** tab appears on the top ribbon.
3. Click **`[About func_SMILES2IMG]`** anytime to verify the installed version, view third-party notices, or check for updates:

<FeatureMedia id="common-about-dialog" />

## 3. Offline / Corporate Network Manual ZIP Install {#offline-zip}

If your workstation blocks direct PowerShell downloads from GitHub:

1. Download `func_SMILES2IMG-x64.zip` (or `func_SMILES2IMG-x86.zip` for 32-bit Office) from [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest).
2. Right-click the `.zip` file ➔ **Properties** ➔ check **Unblock** at the bottom ➔ click **OK**.
3. Extract the archive, open PowerShell inside that folder, and run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

## 4. Security Warnings & Quick Fixes

::: details 🛡️ Windows 11 Smart App Control (SAC) or `#NAME?` After Install
Windows 11 **Smart App Control** can block unsigned open-source Excel `.xll` add-ins when set to strict enforcement. Turning it off **does not disable Microsoft Defender Antivirus**.
1. Open **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security**.
2. Click **App & browser control** ➔ **Smart App Control settings**.
3. Switch it to **Off** and restart Office.
:::

::: details 🔓 Manually Downloaded `.xll` File Blocked by Excel
Right-click the `.xll` file ➔ **Properties** ➔ check **Unblock** ➔ click **OK**, or run the 1-line PowerShell installer above which unblocks all files automatically.
:::

## 5. Updating & Uninstalling

* **Upgrade to the Latest Release:** Click **`[SMILES]` ➔ `[About func_SMILES2IMG]` ➔ `[Check for Updates]`** on the ribbon, or re-run the 1-line PowerShell install command above.
* **Uninstall Cleanly (removes all files & registry entries):**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

## Next Steps

- [**Getting Started: Your First Molecule**](/guide/getting-started) — Try your first formula in Excel, insert a structure in Word, or embed a 3D protein in PowerPoint.
- [**Troubleshooting & FAQ**](/guide/faq) — Solutions for #NAME?, #VALUE!, .FLOAT legacy mode, sharing documents, and SVG export.
