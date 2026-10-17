# Troubleshooting & FAQ

Quick answers to common questions, organized into **Install & Loading**, **Excel Formulas**, **Styling, SVG & SDF Files**, and **Privacy & Updates**.

---

## 1. 🛡️ Install & Loading Issues

### Q. Windows 11 shows "The file format and extension do not match" or Excel formulas show `#NAME?`.
* **Cause:** Windows 11 **Smart App Control (SAC)** blocks unsigned open-source Excel add-ins when strict enforcement is turned on.
* **Rest Assured:** Turning off Smart App Control **does not disable Microsoft Defender Antivirus**—real-time virus protection stays 100% active.
* **10-Second Fix:**
  1. Open Windows **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security**.
  2. Click **App & browser control** ➔ **Smart App Control settings**.
  3. Switch it to **Off** and restart Excel.

### Q. Excel blocks a manually downloaded `.xll` file.
Files downloaded manually from a browser carry a Windows security block flag. Right-click the `.xll` file ➔ **Properties** ➔ check **Unblock** at the bottom ➔ click **OK**, or run the **[10-Second Quick Install](./getting-started)** command which unblocks everything automatically.

---

## 2. 📗 Excel Formula & Display Questions

### Q. Should I use `=SMILES2IMG` or `=SMILES2IMG.FLOAT`?
* **Microsoft 365 or Excel 2024+:** Use `=SMILES2IMG(A2)` for true native in-cell pictures that sort and filter with your table rows.
* **Excel 2016, 2019, 2021:** Older Excel versions do not support native in-cell pictures, so append `.FLOAT`—for example, `=SMILES2IMG.FLOAT(A2)`. This places a cell-locked picture shape that moves and resizes automatically with the cell.

### Q. Why does my formula return `#VALUE!`?
1. Check that the referenced cell contains a valid SMILES string (unclosed ring numbers or parentheses will cause an error).
2. If referencing merged cells, point to the top-left cell of the merged range.
3. **See the Exact Cause:** Click **`[SMILES]`** ➔ **`[Diagnostics]`** on the ribbon to see the exact chemical syntax error message in the diagnostic log.

### Q. Why does the molecule image look small?
Images scale automatically to fit inside the cell without distorting proportions. Simply increase the cell's **row height (e.g., 80–120 pt)** and **column width** to make the structure larger and clearer.

### Q. Can colleagues or Mac users view my documents if they don't have the add-in installed?
Yes! Here is how files behave across platforms:
* **Word (`.docx`) & PowerPoint (`.pptx`):** 2D structures are saved as **standard embedded high-DPI PNG images** and 3D proteins as **standard `.glb` 3D models**. Anyone on Windows, **macOS**, iPad, or Office for the Web can open, view, print, and export your documents to PDF **without installing the add-in**. (When opened later on a PC with the add-in installed, double-click editing still works seamlessly via the embedded Alt Text metadata!)
* **Excel (`.xlsx`):**
  - `.FLOAT` picture shapes and cached cell outputs (MW, Formula, LogP, and Microsoft 365 in-cell images) remain visible on other PCs as long as the cells are not recalculated.
  - **Best Practice for External Sharing:** If a colleague without the add-in triggers a full recalculation (<kbd>F9</kbd>), custom formula cells may show `#NAME?`. Before sending a final spreadsheet externally, either copy the formula columns (<kbd>Ctrl</kbd> + <kbd>C</kbd>) ➔ **Paste as Values**, or export the sheet to **PDF**.

---

## 3. 🎨 Styling, Vector SVG & `.sdf` Files

### Q. How do I switch to ACS Black & White, Dark Mode, or a custom bond color?
* **1-Click Whole-Document Restyle:** In Excel, Word, or PowerPoint, click **`[SMILES]`** ➔ **`[Batch Style]`** on the ribbon to convert all structures in your selection or document to **Publication B&W (ACS style)**, **Dark Mode (White Ink)**, **Standard Color**, or **White Background** in one click.
* **Per-Formula Styling in Excel (Including Custom Ink Color):**
  - Black & White line art: `=SMILES2IMG(A2, "", "bw")`
  - Dark Mode: `=SMILES2IMG(A2, "#1a1a1a", "white")`
  - Custom bond/atom ink color: `=SMILES2IMG(A2, "", "ink=#1e40af")`

### Q. Can I copy lossless vector graphics (`.svg`) for Adobe Illustrator or ChemDraw?
Yes! Select any molecule cell or picture and click **`[SMILES]`** ➔ **`[Copy SVG]`** on the ribbon (or right-click ➔ **`Chemical Structure`** ➔ **`Copy Vector Graphic (SVG)`**). Paste directly into Adobe Illustrator, Inkscape, or vector editors with zero loss of quality.

### Q. Can I open or save `.sdf` chemical database files in Excel?
Yes! Click **`[SMILES]`** ➔ **`[SDF File]`** on the Excel ribbon:
* **`Import SDF File (.sdf)...`:** Reads any multi-compound `.sdf` file and generates an 8-column medicinal chemistry worksheet (`Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `Lipinski`) powered by live Excel formulas.
* **`Export Worksheet to SDF...`:** Exports your worksheet compounds back to a standard V2000 `.sdf` file.

### Q. How do I copy Molecular Weight, Formula, LogP, and Lipinski info into a table at once?
**Right-click** any molecule ➔ **`Chemical Structure`** ➔ **`Copy Molecule Info`** and choose your format:
* **`Copy All Summary (Text)`:** Copies a **2-row table (Headers + Data)**—ideal for starting a new table in Excel or Word.
* **`Copy as Excel Row (TSV Columns)`:** Copies **only the 12 data columns (no header)**—ideal for appending rows into an existing table.
* **`Copy as Multi-Line List`:** Ready to paste as bullet points into PowerPoint slide text boxes.
* **`Copy as Markdown Table` / JSON:** Great for Notion, GitHub, and electronic lab notebooks (also available inside the **Detailed Molecule Info Window**).

---

## 4. 🔒 Privacy (Offline Mode) & Updates

### Q. Are my chemical structures ever sent to the cloud or external servers?
**Never. Everything runs 100% locally and offline on your PC.**  
The chemical rendering engine (EPAM Indigo), 2D sketcher (Ketcher), and 3D protein viewer (Mol\*) are all bundled locally inside the add-in. Proprietary drug candidates and unpublished structures never leave your computer. *(The only optional web request is looking up an uncached common name from PubChem when using `=NAME2SMILES` or `=SMILES2NAME`, and results are cached locally on your PC).*

### Q. How do I check for updates or upgrade?
Click **`[SMILES]`** ➔ **`[About SMILES]`** on the ribbon and click **`[Check for Updates]`**, or re-run the one-line PowerShell install command at any time to upgrade to the latest release.

### Q. How does `func_SMILES2IMG` compare to ChemDraw for Excel, JChem for Office, or Chem4Word?
| Feature | **func_SMILES2IMG (This Suite)** | **ChemDraw / JChem for Office** | **Chem4Word** |
| :--- | :--- | :--- | :--- |
| **License & Cost** | **100% Free & Open-Source (MIT)** | Paid commercial subscription | Free & open-source |
| **Supported Office Apps** | **Excel + Word + PowerPoint** | Excel + Word + PowerPoint | Word only |
| **Excel Structure Cells** | **Native in-cell formulas (`=SMILES2IMG`)** & `.FLOAT` shapes (no special sheet conversion needed) | Requires special worksheet conversion or proprietary objects | N/A |
| **3D Macromolecules** | **Built-in Mol\* (rotatable `.glb` 3D models & 2D PNG)** | Requires separate 3D software | Not supported |
| **Installation Rights** | **No Admin rights needed** (10-second per-user setup) | Requires Admin installer & license activation | Requires setup wizard |

---

## 🔗 Related Guides

<div class="feature-grid-2">
  <div class="feature-card">
    <h4>🚀 <a href="/guide/getting-started">Install & 1-Minute Quickstart</a></h4>
    <p>One-line PowerShell installer and offline manual setup instructions for corporate networks.</p>
  </div>
  <div class="feature-card">
    <h4>🛠️ <a href="/tools/formula-builder">Formula Generator & Style Simulator</a></h4>
    <p>Click options to preview rendered structures in real time and copy the ready-to-use Excel formula.</p>
  </div>
</div>


