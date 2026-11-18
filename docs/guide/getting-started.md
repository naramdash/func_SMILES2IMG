# 1-Min Install & Your First Molecule

Install **`func_SMILES2IMG`** in 10 seconds without administrator rights, then jump straight to your Office app guide (**Excel Recipe Cookbook**, **Word 6-Chapter Manuscript Tutorial**, or **PowerPoint Visual Slide Gallery**).

---

## Step 1. 10-Second Auto-Install (PowerShell)

<QuickInstallBox />

::: tip 🚀 What's Included in `v1.8.0`
Running the command above automatically installs **`func_SMILES2IMG v1.8.0`**, featuring unified `Chemical Structure >` cascading right-click menus across Excel/Word/PowerPoint, cross-host text/name parity & aspect-ratio lock, refreshed Windows Info & About dialogs with 2D preview, `=SMILES2HEAVY`–`=SMILES2CHIRAL` 3D complexity formulas, and 12-column `.sdf` import.
:::

---

## Step 2. Restart Office & Verify the `[SMILES]` Ribbon Tab

1. **Completely close and reopen** Excel, Word, or PowerPoint.
2. Confirm that **`[SMILES]`** appears on the top Office ribbon, and click **`[About func_SMILES2IMG]`** anytime to verify the installed version or check for updates:

<FeatureMedia id="common-about-dialog" />

---

## Step 3. Pick Your Office App & Render Your First Structure

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="./excel" class="s2i-card">
    <div class="s2i-card-title">📗 Microsoft Excel — Recipe Cookbook</div>
    <p class="s2i-card-desc">Result-first recipes for structure tables, Name/CAS lookup, Lipinski/SAR, and 40 formulas</p>
  </a>
  <a href="./word" class="s2i-card">
    <div class="s2i-card-title">📘 Microsoft Word — 6-Chapter Manuscript Tutorial</div>
    <p class="s2i-card-desc">Build one complete research manuscript from prose figures and tables to ACS B&amp;W submission</p>
  </a>
  <a href="./powerpoint" class="s2i-card">
    <div class="s2i-card-title">📙 Microsoft PowerPoint — Visual Slide Gallery</div>
    <p class="s2i-card-desc">Browse dark-mode structure cards, synthesis slides, and 360° rotatable 3D proteins</p>
  </a>
</div>

### 📗 Option A. Microsoft Excel (Formula or `Ctrl + Shift + K`)
Paste either formula below into any empty cell (on Excel 2016–2021, append **`.FLOAT`**, e.g. `=NAME2IMG.FLOAT("Aspirin")`):

```text
=NAME2IMG("Aspirin")
```
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```

<FeatureMedia id="excel-overview" />

::: tip 📥 Download Live Interactive 8-Tab Master Excel Showcase
Explore all capabilities without typing formulas from scratch! Download the official **8-tab master guide workbook**:
<div style="margin: 10px 0 4px 0;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card" style="display: block; padding: 12px 16px; border: 1.5px solid var(--vp-c-brand-1); border-radius: 8px; text-decoration: none;">
    <div style="font-weight: 700; color: var(--vp-c-brand-1); font-size: 1rem;">📥 Download SMILES2IMG_Comprehensive_Showcase.xlsx (.xlsx)</div>
    <div style="color: var(--vp-c-text-2); font-size: 0.88rem; margin-top: 4px;">0. Quick Guide · 1. Name/CAS/SMILES · 2. Lipinski &amp; 3D Complexity · 3. Substructure &amp; Similarity · 4. Styles Matrix · 5. Reactions &amp; KET · 6. Legacy .FLOAT · 7. 3D &amp; Ribbon Tools</div>
  </a>
</div>
:::

### 📘 Option B. Microsoft Word (Prose & Summary Tables)
Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>), sketch in Ketcher, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert inline.

<FeatureMedia id="word-overview" />

### 📙 Option C. Microsoft PowerPoint (Dark Mode & 3D Slides)
Insert transparent structures for dark slides, sketch reactions, or embed 360°-rotatable 3D proteins (`Mol*`).

<FeatureMedia id="powerpoint-overview" />

---

::: details ❓ Seeing `#NAME?` in Excel or Missing the `[SMILES]` Tab? (30-Second Fix)
1. **Did you restart Excel / Word / PowerPoint?**  
   Add-ins load on startup. Close all open Office windows completely and reopen the app.
2. **Are you on Excel 2016, 2019, or 2021?**  
   In-cell image formulas (`=SMILES2IMG`) require Microsoft 365 or Excel 2024+. On Excel 2016–2021, append **`.FLOAT`**: `=NAME2IMG.FLOAT("Aspirin")`.
3. **Windows 11 Smart App Control (SAC):**  
   Open **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security** ➔ **App & browser control** ➔ **Smart App Control settings** ➔ **Off**, then restart Office. *(Microsoft Defender Antivirus remains 100% active.)*
:::

---

## 🔄 Updating & Uninstalling

* **Upgrade to the Latest Release:** Click **`[SMILES]` ➔ `[About func_SMILES2IMG]` ➔ `[Check for Updates]`** on the ribbon, or re-run the 1-line PowerShell installer above.
* **Uninstall Cleanly:**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

---

## 📚 Corporate Network & System Info

::: details 🏢 Offline / Corporate Network Manual ZIP Installation
1. Download `func_SMILES2IMG-x64.zip` (or `func_SMILES2IMG-x86.zip`) from [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest).
2. Right-click the `.zip` file ➔ **Properties** ➔ check **Unblock** ➔ click **OK**.
3. Extract the archive, open PowerShell in that folder, and run `powershell -ExecutionPolicy Bypass -File install.ps1`.
:::

::: details 💻 System Requirements
| Component | Requirement | Notes |
| :--- | :--- | :--- |
| **OS** | Windows 10 / Windows 11 | Supports both 32-bit and 64-bit systems |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | Desktop installations |
| **Runtime** | .NET Framework 4.8 & Edge WebView2 | Pre-installed on Windows 10 & Windows 11 |
:::
