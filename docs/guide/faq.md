# Frequently Asked Questions (FAQ) & Troubleshooting

Find quick answers to common questions and troubleshooting tips for **SMILES for Microsoft Office Suite**.

---

## ❓ General & Excel Questions

### Q. Which formula should I use for my Excel version?
* **Microsoft 365 or Excel 2024+:** Use `=SMILES2IMG(...)` for real in-cell pictures that sit directly inside cell contents and move/scale automatically.
* **Excel 2016, 2019, 2021:** Use `=SMILES2IMG.FLOAT(...)` because older Excel builds lack native in-cell picture support. Floating images automatically move and size with cells (`xlMoveAndSize`).

### Q. The formula returns `#VALUE!`. How do I find why it failed?
1. Verify that the referenced cell contains a valid SMILES string.
2. If referencing merged cells, make sure you point to the top-left cell of the merge.
3. **Inspect Exact Chemical Errors:** From the Excel ribbon, click **[SMILES] ➔ [Show diagnostics]**. The diagnostic log displays the exact chemical syntax error (such as unclosed rings `cycle not closed`, invalid valences, or unrecognized atoms) and character position.

### Q. The molecular image looks too small.
Images scale automatically to preserve aspect ratios without clipping. Increase the cell's row height (e.g., 80–120 pt) and column width (e.g., 25–40) to allow a larger, clear rendering.

### Q. How do I match monochrome journals or dark mode themes?
* **Black & White Line Art:** Pass `"bw"` as the 3rd argument for 100% monochrome patent and ACS print style (`=SMILES2IMG(A2, "", "bw")`).
* **Dark Mode:** Pass a dark background (e.g. `"#1a1a1a"`) as the 2nd argument and white ink (`"white"`) as the 3rd argument:
  ```text
  =SMILES2IMG(A2, "#1a1a1a", "white")
  ```
* **1-Click Whole-Document Normalization (`[Batch Style]`):** In Excel, Word, or PowerPoint, click **`[SMILES] ➔ [Batch Style]`** on the ribbon to convert all structures in the selection, active worksheet/slide, or entire document to **ACS Standard (B&W)**, **Dark Mode (White Ink, Transparent)**, **Color (Transparent)**, or **Color (White BG)** in one click.

### Q. How do I export crisp vector graphics (`.svg`) for Adobe Illustrator, Inkscape, or LaTeX?
Select any chemical cell or structure picture in Excel, Word, or PowerPoint and click **`[SMILES] ➔ [Copy SVG]`** (or right-click ➔ **`Chemical Structure ➔ Copy Vector Graphic (SVG)`**). Both the `image/svg+xml` MIME clipboard stream and raw `<svg>` XML text are placed onto the Windows clipboard for lossless pasting.

### Q. Can I import or export multi-molecule `.sdf` files in Excel?
Yes! On the Excel **`[SMILES]`** ribbon tab, click **`[SDF File]`** in the **Data Exchange** group:
* **Import `.sdf` File...:** Parses multi-record `.sdf` files into a clean worksheet table with automatic `=SMILES2IMG` pictures, SMILES, Molecular Formula, Molecular Weight, and all SD `<TAG>` properties.
* **Export `.sdf` File...:** Exports selected rows or the active table of SMILES strings and adjacent columns into a standard V2000 `.sdf` file.

### Q. Are my chemical structures sent to the cloud or external servers?
**No. 100% strictly local and offline.** All chemical engines (EPAM Indigo C++, EPAM Ketcher 2D, and Mol* 3D) are packaged locally inside the add-in. Structures and reaction graphs are computed entirely in-memory on your CPU/GPU without cloud telemetry, making it completely safe for proprietary drug candidates, patents, and confidential corporate research. *(The only optional web call is PubChem common name lookup for unknown SMILES, which can also be pre-cached).*

### Q. How do I use the new Molecule Info Dialog and export multi-format data?
Right-click any chemical cell or picture and select **Chemical Structure ➔ Copy Molecule Info ➔ Detailed Molecule Info Window...** (or click **`[Molecule Info]`** on the ribbon).
* Inspect all chemical properties in a clean **DataGridView table** featuring custom high-DPI vector chemical icons for each metric.
* Use the **Export Split Button** at the bottom to copy properties in one click as:
  - **Excel Row (TSV):** Tab-separated columns that paste across adjacent Excel cells or Word/PowerPoint table rows.
  - **Markdown Table:** GitHub-flavored markdown table for lab notebooks and documentation.
  - **Multi-Line List:** Bulleted property list for PowerPoint slides and notes.
  - **One-Line Text:** Compact text summary.
  - **JSON:** Formatted JSON data payload for computational chemistry workflows.

### Q. How do I check for updates or upgrade to the latest version?
Click **`[SMILES] ➔ [About SMILES]`** in Excel, Word, or PowerPoint, then click **`[Check for Updates]`**. If a newer release is available on GitHub, confirm the prompt to launch the automated one-click PowerShell updater.

---

## 🛡️ Security & Loading Issues

### Q. [Smart App Control] Excel displays "The file format and extension do not match" or formulas evaluate to `#NAME?`.
* **Cause:** Windows 11 **Smart App Control (SAC)** blocks unsigned open-source XLL add-in binaries when enforcement mode is active.
* **Important:** Disabling Smart App Control does **NOT** disable your antivirus; **real-time Microsoft Defender protection remains 100% active**.
* **Fix (10 seconds):**
  1. Open Windows **Settings** (`Win + I`) ➔ **Privacy & security** ➔ **Windows Security**.
  2. Click **App & browser control** ➔ **Smart App Control settings**.
  3. Set to **Off** and restart Excel.

### Q. Excel blocks or refuses to load the XLL add-in.
Web downloads are marked by Windows with a `Zone.Identifier` stream. Right-click the `.xll` file, select **Properties**, check **Unblock** at the bottom, and click **OK**.
Alternatively, run the [One-Line Quick Install](./getting-started) command, which removes security blocks automatically.
