# Frequently Asked Questions (FAQ)

Answers to common workflow questions about styling, vector SVG copy, `.sdf` files, property tables, and updates. Looking for a specific error fix? See [Installation Blocks](/guide/troubleshooting/install), [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab), [Document Sharing](/guide/troubleshooting/document-sharing), or app-specific troubleshooting for [Excel](/excel/troubleshooting/formula-errors), [Word](/word/troubleshooting/double-click), [PowerPoint](/powerpoint/troubleshooting/dark-slides), and [Editor](/editor/troubleshooting/webview2-loading).

## 1. Styling & Vector Graphics

### Q. How do I switch to ACS Black & White, Dark Mode, or a custom bond color?
* **Whole-Document Restyle:** In Excel, Word, or PowerPoint, click **`[SMILES]` ➔ `[Batch Style]`** on the ribbon to convert structures in your selection or document to **Publication B&W (ACS style)**, **Dark Mode (White Ink)**, **Standard Color**, or **White Background**.
* **Per-Formula Styling in Excel (Including Custom Ink Color):**
  - Black & White line art: `=SMILES2IMG(A2, "", "bw")`
  - Dark Mode (White Ink + Transparent BG): `=SMILES2IMG(A2, "trans", "white")`
  - Custom bond/atom ink color: `=SMILES2IMG(A2, "", "ink=#1e40af")`

### Q. Can I copy lossless vector graphics (`.svg`) for Adobe Illustrator or Inkscape?
Yes! Select any molecule cell or picture and click **`[SMILES]` ➔ `[Copy SVG]`** on the ribbon (or right-click ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**). Paste directly into Adobe Illustrator, Inkscape, or vector editors without pixelation.

## 2. `.sdf` Files & Molecular Property Export

### Q. Can I open `.sdf` chemical database files in Excel?
Yes! Click **`[SMILES]` ➔ `[Import SDF]`** on the Excel ribbon:
* **`Import SDF (.sdf)...`:** Reads any multi-compound `.sdf` file and generates an 11-column medicinal chemistry worksheet (`Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `TPSA`, `Rings`, `Aromatic Rings`, `Fsp3`) powered by live Excel formulas.

### Q. How do I copy Molecular Weight, Formula, LogP, and TPSA info into a table at once?
**Right-click** any molecule ➔ **`Chemical Structure` ➔ `Copy Molecule Info`** and choose your format:
* **`Copy All Summary (Text)`:** Copies a **2-row table (Headers + Data)**—ideal for starting a new table in Excel or Word.
* **`Copy as Excel Row (TSV Columns)`:** Copies **only the 12 data columns (no header)**—ideal for appending rows into an existing table.
* **`Copy as Multi-Line List`:** Ready to paste as bullet points into PowerPoint slide text boxes.
* **`Copy as Markdown Table` / JSON:** Great for Notion, GitHub, and electronic lab notebooks (inside **`Detailed Molecule Info Window...`**, choose a format from the bottom dropdown and click **`Copy All`**, or double-click any individual property row to copy it).

## 3. Privacy, Updates & Compatibility

### Q. Are my chemical structures sent to external servers?
**No — core 2D rendering, 2D sketching, and all property/SAR calculations run locally on your PC.** Internet access is only used optionally for uncached PubChem name/CAS lookups, RCSB PDB ID fetches in Mol\*, and update checks. See [Local-First Privacy & Sharing](/guide/#privacy) for the complete breakdown.

### Q. Should I use `=SMILES2IMG` or `=SMILES2IMG.FLOAT`?
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(A2)` for true native in-cell pictures that sort and filter with table rows.
* **Excel 2016 / 2019 / 2021:** Append `.FLOAT` (`=SMILES2IMG.FLOAT(A2)`) to create cell-anchored floating shapes (`xlMoveAndSize`).

### Q. How do I check for updates or upgrade?
Click **`[SMILES]` ➔ `[About func_SMILES2IMG]`** on the ribbon and press **`[Check for Updates]`**, or re-run the [PowerShell Install](/guide/install) command anytime to upgrade to the latest release:

<FeatureMedia id="common-about-dialog" />

## Related Guides

- [**Installation & Setup**](/guide/install) — One-line PowerShell installer, offline manual ZIP setup, and update/uninstall commands.
- [**All 49 Excel Functions**](/excel/formulas) — Full parameter catalog, return data types, and shared styling tokens.
