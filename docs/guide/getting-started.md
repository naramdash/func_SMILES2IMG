# 1-Min Install & Your First Molecule

Install **`func_SMILES2IMG`** in 10 seconds and render your first 2D chemical structure or 3D protein inside Excel, Word, or PowerPoint.

---

## Step 1. 10-Second Auto-Install (PowerShell)

<QuickInstallBox :hide-quickstart-btn="true" />

---

## Step 2. Restart Office & Check the `[SMILES]` Ribbon Tab

1. **Completely close and reopen** Excel, Word, or PowerPoint.
2. Confirm that the **`[SMILES]`** tab appears on the top Office ribbon:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Office" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## Step 3. Render Your First Molecule (Pick Your App)

<p align="center">
  <img src="/hero-demo.gif" alt="func_SMILES2IMG 1-Minute Quickstart Demo" width="100%" style="max-width: 760px; border-radius: 10px; box-shadow: 0 6px 20px rgba(0,0,0,0.12);" />
</p>

### 📗 Option A. Using Microsoft Excel? (Copy & Paste a Formula)
Copy either formula below and paste it into any empty cell (widen the row height slightly for a larger image):

```text
=NAME2IMG("Aspirin")
```
Or from a SMILES string:
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```
> 💡 **Using Excel 2016, 2019, or 2021?**  
> Append **`.FLOAT`** to the formula name—for example, `=NAME2IMG.FLOAT("Aspirin")` or `=SMILES2IMG.FLOAT(A2)`!

### 📘📙 Option B. Using Word or PowerPoint? (Press `Ctrl + Shift + K`)
Word and PowerPoint do not use cell formulas—use the ribbon button or shortcut instead:
1. Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** on the top ribbon, or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>.
2. Sketch your molecule or reaction scheme inside the built-in **Ketcher 2D Editor**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**) to insert a crisp structure directly onto your page or slide!
4. **Edit Anytime:** **Double-click** any inserted molecule picture to reopen the editor and modify it in place.

---

::: details ❓ Seeing `#NAME?` in Excel or Missing the `[SMILES]` Tab? (30-Second Fix)
If formulas show `#NAME?` or the `[SMILES]` tab is missing right after installation, check these 3 common causes:

1. **Did you restart Excel / Word / PowerPoint?**  
   Add-ins load when Office starts. Close all open Office windows completely and reopen the app.
2. **Are you on Excel 2016, 2019, or 2021?**  
   In-cell image formulas (`=SMILES2IMG`) require Microsoft 365 or Excel 2024+. On Excel 2016–2021, use **`.FLOAT`** formulas instead:
   ```text
   =NAME2IMG.FLOAT("Aspirin")
   ```
3. **Windows 11 Smart App Control (SAC) Blocking Unsigned Open-Source Add-ins:**  
   If Windows 11 shows a *"file format and extension do not match"* warning:
   - Open Windows **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security** ➔ **App & browser control**.
   - Click **Smart App Control settings** and switch it to **Off**, then restart Office. *(Microsoft Defender Antivirus real-time protection remains 100% active.)*
:::

---

## 🔄 Updating & Uninstalling

* **Upgrade to the Latest Release:**
  - Click **`[SMILES]` ➔ `[About SMILES]` ➔ `[Check for Updates]`** on the Office ribbon, or re-run the 1-line PowerShell install command anytime to upgrade automatically.
* **Uninstall Cleanly:**
  - Run this single command in PowerShell to remove all add-in files and registry entries:
    ```powershell
    irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
    ```

---

## 📚 Corporate Network & System Info (Click to Expand)

::: details 🏢 Offline / Corporate Network Manual ZIP Installation
If your company firewall blocks GitHub scripts:
1. Download `func_SMILES2IMG-x64.zip` (or `func_SMILES2IMG-x86.zip` for 32-bit Office) from [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest).
2. Right-click the `.zip` file ➔ **Properties** ➔ check **Unblock** at the bottom ➔ click **OK**.
3. Extract the archive, open PowerShell in that folder, and run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```
:::

::: details 💻 System Requirements
| Component | Requirement | Notes |
| :--- | :--- | :--- |
| **OS** | Windows 10 / Windows 11 | Supports both 32-bit and 64-bit systems |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | Desktop installations |
| **Runtime** | .NET Framework 4.8 & Edge WebView2 | Pre-installed on Windows 10 & Windows 11 |
:::
