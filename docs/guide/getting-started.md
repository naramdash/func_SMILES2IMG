# Install & 1-Minute Quickstart

**func_SMILES2IMG (SMILES for Office)** is a free, 100% offline add-in suite for Microsoft Excel, Word, and PowerPoint that lets you draw 2D chemical structures and reactions, calculate molecular properties via formulas, and visualize 3D proteins.

---

## Step 1. 10-Second Install

1. **Close Excel, Word, and PowerPoint** if they are currently open.
2. Open **PowerShell** from the Windows Start menu and paste this command (no Administrator rights required):

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

3. Once you see `"Installation complete!"`, you are ready to go!

---

## Step 2. Your First Molecule in 1 Minute

Open Excel, Word, or PowerPoint and try these 3 core workflows right away:

### ① Render a Molecule by Formula in Excel
Open Excel and type either formula into any empty cell (widen the row height slightly for a larger view):

```text
=NAME2IMG("Aspirin")
```
Or from a SMILES string:
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```
> 💡 **Using Excel 2016, 2019, or 2021?**  
> Append `.FLOAT` to the formula name—for example, `=NAME2IMG.FLOAT("Aspirin")` or `=SMILES2IMG.FLOAT(A2)`—to place a cell-locked picture shape that moves and resizes with the cell!

### ② Draw a Molecule or Reaction by Hand (`Ctrl + Shift + K`)
In Excel, Word, or PowerPoint:
1. Click the **`[SMILES]`** tab on the top ribbon and click **`[Insert / Edit (SMILES)]`** in Excel (or **`[Insert / Edit Structure]`** in Word & PowerPoint)—or simply press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>.
2. Sketch your molecule or reaction scheme inside the built-in **Ketcher 2D Editor**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**) to insert a crisp structure directly into your cell, document, or slide!

### ③ Double-Click to Edit & Right-Click to Copy Properties
* **Edit In-Place:** **Double-click** any inserted molecule cell or picture to reopen the sketcher and modify it in place.
* **Copy Formula & MW:** **Right-click** any molecule ➔ **`Chemical Structure`** ➔ **`Copy Molecule Info`** to copy Molecular Weight (`180.16 g/mol`), Formula (`C9H8O4`), InChIKey, or a full table row in one click.

---

## Step 3. Explore by Application

Pick the guide for the Office app you are working in:

* 📗 **[Microsoft Excel Guide](/guide/excel):** Build compound tables, auto-calculate MW & Lipinski Ro5, highlight substructures, and import/export `.sdf` files
* 📘 **[Microsoft Word Guide](/guide/word):** Insert inline structures inside paragraphs and tables, and convert entire manuscripts to ACS Black & White
* 📙 **[Microsoft PowerPoint Guide](/guide/powerpoint):** Use transparent backgrounds for dark slides and embed 360°-rotatable 3D protein models
* 📐 **[Excel Formula Cheat Sheet](/reference/formula):** Quick lookup table for all 25+ chemical functions and styling options

---

## 🛠️ Troubleshooting Install Issues (10-Second Fix)

### Windows 11 Warning or `#NAME?` in Excel (Smart App Control)
If Windows 11 **Smart App Control (SAC)** is enabled in strict mode, Windows may block unsigned open-source add-ins with a *"file format and extension do not match"* message or show `#NAME?` in Excel.

::: tip Your Antivirus Stays 100% Active
Turning off Smart App Control **does not disable Microsoft Defender Antivirus**. Real-time virus and threat protection remains 100% active.
:::

1. Open Windows **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security**.
2. Click **App & browser control** ➔ **Smart App Control settings**.
3. Switch it to **Off**, then restart Excel / Word / PowerPoint.

---

## 🔄 Updating & Uninstalling

* **Check for Updates / Upgrade:**
  - Click **`[SMILES]`** on the Office ribbon ➔ **`[About SMILES]`** ➔ **`[Check for Updates]`**.
  - Or re-run the one-line PowerShell install command anytime to upgrade to the latest release (`v1.7.0`).
* **Uninstall Cleanly:**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

---

## 📚 Advanced & Offline Setup (Click to Expand)

::: details 🏢 Offline / Corporate Network Manual Installation
If your organization blocks remote PowerShell scripts:
1. Download `Smiles2Img-Office-x64.zip` (or `x86.zip` for 32-bit Office) and `install.ps1` from [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases).
2. Right-click the `.zip` file ➔ **Properties** ➔ check **Unblock** at the bottom ➔ click **OK**.
3. Extract the archive to `%APPDATA%\Microsoft\AddIns\Smiles2Img`.
4. Open PowerShell in that folder and run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```
:::

::: details 💻 System Requirements
| Component | Requirement | Notes |
| :--- | :--- | :--- |
| **OS** | Windows 10 / Windows 11 | Supports both 32-bit and 64-bit systems |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | Desktop installations |
| **Runtime** | .NET Framework 4.8 & Edge WebView2 | Pre-installed on modern Windows 10 & Windows 11 |
:::

::: details ⚙️ What the Installer Does Under the Hood
1. Detects whether your Office installation is 32-bit (`x86`) or 64-bit (`x64`).
2. Downloads the latest release package to `%APPDATA%\Microsoft\AddIns\Smiles2Img`.
3. Strips Windows download block flags (`Unblock-File`).
4. Registers the Excel XLL add-in and Word/PowerPoint COM add-ins under your current user registry (`HKCU`) without requiring admin rights.
:::

::: details ✨ What's New in v1.7.0
* **In-Cell SAR & Drug-Likeness:** Screen compounds with `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2LIPINSKI`, `=SMILES2LOGP`, and `=SMILES2TPSA`.
* **Chemical Name & CAS Lookup:** Render structures directly from names or CAS numbers via `=NAME2IMG("Aspirin")` and `=NAME2SMILES("50-78-2")`.
* **1-Click Batch Style:** Unify all structures in a sheet, document, or slide deck into Standard Color, ACS Publication B&W, White Background, or Dark Mode.
* **Vector SVG Copy & Excel `.sdf` I/O:** Copy lossless SVG vectors for Illustrator/Inkscape and import/export `.sdf` chemical tables in Excel.
* **Redesigned Molecule Info Window:** Clean property table with one-click export to Excel rows (TSV), Markdown tables, bullet lists, and JSON.
:::
