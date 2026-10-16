# Getting Started & Installation

Welcome to **SMILES for Microsoft Office Suite** — an all-in-one chemical structure rendering, 3D macromolecular visualization (Mol*), molecular property calculation, and visual editing solution for Microsoft Excel, Microsoft Word, and Microsoft PowerPoint.

---

## ⚡ Quick Install (Recommended)

Open **PowerShell** (no Administrator privileges required for standard user installs) and run:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 Preview Release (`v1.7.0-preview.5`)
To test the latest preview release featuring DataGridView Molecule Info Table, Multi-Format Molecule Summary (Excel TSV Row, Multi-Line, Markdown Table, JSON, Split Button), SAR Substructure/Similarity (`=SMILES.MATCH`, `=SMILES.SIMILARITY`), Lipinski/LogP/TPSA, Reverse Name/CAS lookup (`=NAME2IMG`), Vector SVG Copy, Batch Style Normalizer, and Excel SDF Import/Export:
```powershell
$v='v1.7.0-preview.5'; irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```
:::

### What the Automated Installer Does
1. **Detects Office Environment:** Automatically checks whether your Office installation is 32-bit (`x86`) or 64-bit (`x64`) and detects installed apps (Excel, Word, PowerPoint).
2. **Downloads Latest Binaries:** Fetches the unified release suite from GitHub Releases.
3. **Places in Safe Add-in Directory:** Installs to `%APPDATA%\Microsoft\AddIns\Smiles2Img`.
4. **Removes Security Flags:** Runs `Unblock-File` to strip Windows download block flags (`Zone.Identifier`).
5. **Registers Add-ins:**
   - Registers Excel XLL add-in under `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN`.
   - Registers Word COM Add-in (`Smiles2Img.WordAddin`) in `HKCU\Software\Microsoft\Office\Word\Addins`.
   - Registers PowerPoint COM Add-in (`Smiles2Img.PowerPointAddin`) in `HKCU\Software\Microsoft\Office\PowerPoint\Addins`.

---

## 📊 Runtime Requirements

| Component | Required Version | Notes |
| :--- | :--- | :--- |
| **Operating System** | Windows 10 / Windows 11 | 32-bit & 64-bit architectures supported |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | Desktop versions (Click-to-Run or MSI) |
| **.NET Framework** | .NET Framework 4.8 or newer | Pre-installed on Windows 10 (1903+) & Windows 11 |
| **Microsoft Edge WebView2** | Evergreen Runtime | Powers the offline Ketcher 2D editor and Mol* 3D macromolecular viewer (pre-installed on Windows 10/11) |

---

## 🛡️ Windows 11 Smart App Control

If your Windows 11 system has **Smart App Control (SAC)** enabled in strict enforcement mode, you may see a notification stating *"The file format and extension do not match"* or Excel formulas evaluating to `#NAME?`.

::: tip Rest Assured
Turning off Smart App Control **never disables Microsoft Defender Antivirus**. Real-time cloud-delivered virus and threat protection remains **100% active**.
:::

### Solution (Takes 10 Seconds)
1. Open Windows **Settings** (`Win + I`) ➔ **Privacy & security** ➔ **Windows Security**.
2. Click **App & browser control** ➔ **Smart App Control settings**.
3. Toggle the setting to **Off**.
4. Restart Excel / Word / PowerPoint. The add-in loads immediately.

---

## 🖐️ Manual Installation

If your IT policies restrict running remote PowerShell web scripts:

1. Download the latest `Smiles2Img-Office-x64.zip` (or `Smiles2Img-Office-x86.zip`) and `install.ps1` from [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases).
2. Right-click the downloaded `.zip` file ➔ **Properties** ➔ check **Unblock** ➔ click **OK**.
3. Extract the contents to `%APPDATA%\Microsoft\AddIns\Smiles2Img`.
4. Run `install.ps1` locally inside PowerShell:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

---

## 🔄 Checking for Updates & Upgrading

* **In-App Update Checker:** In Excel, Word, or PowerPoint, open the **`[SMILES]`** ribbon tab, click **`[About SMILES]`**, and click **`[Check for Updates]`**. If a newer release is available on GitHub, the dialog offers a one-click copy of the upgrade command and opens the release page.
* **Command-Line Upgrade:** Running the one-line `install.ps1` command again at any time automatically replaces existing binaries with the newest release.

---

## 🗑️ Uninstallation

To cleanly remove all add-in registrations and cached files:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
```
Or run `uninstall.ps1` locally from `%APPDATA%\Microsoft\AddIns\Smiles2Img`.
