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
3. **See the Exact Cause:** Click **`[SMILES] ➔ [Settings & Diagnostics]`** on the ribbon to see the exact chemical syntax error message in the log.

### Q. Why does the molecule image look small?
Images scale automatically to fit inside the cell without distorting proportions. Simply increase the cell's **row height (e.g., 80–120 pt)** and **column width** to make the structure larger and clearer.

---

## 3. 🎨 Styling, Vector SVG & `.sdf` Files

### Q. How do I switch to ACS Black & White or Dark Mode?
* **1-Click Whole-Document Restyle:** In Excel, Word, or PowerPoint, click **`[SMILES] ➔ [Batch Style]`** on the ribbon to convert all structures in your selection or document to **Publication B&W (ACS style)**, **Dark Mode (White Ink)**, **Standard Color**, or **White Background** in one click.
* **Per-Formula Styling in Excel:**
  - Black & White line art: `=SMILES2IMG(A2, "", "bw")`
  - Dark Mode: `=SMILES2IMG(A2, "#1a1a1a", "white")`

### Q. Can I copy lossless vector graphics (`.svg`) for Adobe Illustrator or ChemDraw?
Yes! Select any molecule cell or picture and click **`[SMILES] ➔ [Copy SVG]`** on the ribbon (or right-click ➔ **`Chemical Structure ➔ Copy Vector Graphic (SVG)`**). Paste directly into Adobe Illustrator, Inkscape, or vector editors with zero loss of quality.

### Q. Can I open or save `.sdf` chemical database files in Excel?
Yes! Click **`[SMILES] ➔ [SDF File]`** on the Excel ribbon:
* **Import `.sdf` File...:** Reads any multi-compound `.sdf` file and creates a formatted worksheet table with rendered structures (`=SMILES2IMG`), SMILES, Formula, Molecular Weight, and all SD property columns.
* **Export `.sdf` File...:** Exports your selected compound rows back to a standard V2000 `.sdf` file.

### Q. How do I copy Molecular Weight, Formula, LogP, and Lipinski info into a table at once?
**Right-click** any molecule ➔ **`Chemical Structure ➔ Copy Molecule Info`** and choose your format:
* **Copy as Excel Row (TSV Columns):** Pastes cleanly across adjacent Excel columns or directly into Word/PowerPoint table cells.
* **Copy as Multi-Line List:** Ready to paste as bullet points into PowerPoint slide text boxes.
* **Copy as Markdown Table / JSON:** Great for Notion, GitHub, and electronic lab notebooks (also available inside the **Detailed Molecule Info Window**).

---

## 4. 🔒 Privacy (Offline Mode) & Updates

### Q. Are my chemical structures ever sent to the cloud or external servers?
**Never. Everything runs 100% locally and offline on your PC.**  
The chemical rendering engine (EPAM Indigo), 2D sketcher (Ketcher), and 3D protein viewer (Mol\*) are all bundled locally inside the add-in. Proprietary drug candidates and unpublished structures never leave your computer. *(The only optional web request is looking up a common name from PubChem when using `=SMILES2NAME` on an uncached molecule).*

### Q. How do I check for updates or upgrade?
Click **`[SMILES] ➔ [About SMILES]`** on the ribbon and click **`[Check for Updates]`**, or re-run the one-line PowerShell install command at any time to upgrade to the latest release.
