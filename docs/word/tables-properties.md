# 2. SAR Tables & Experimental Properties

Build manuscript SAR and compound summary tables that never collapse column widths or clip structure tops—and copy locally calculated Chemical Formula, Molecular Weight, HRMS Exact Mass, LogP, and TPSA data with zero manual typing errors.

<FeatureMedia id="word-ch3-compound-table" />

## 1. Three Rules for Bulletproof Word Compound Tables

Word automatically resizes columns and clips tall inline graphics unless you configure these three table settings:

1. **Lock Column Widths:** Right-click the table ➔ **AutoFit** ➔ **Fixed Column Width**. Even when you double-click a cell structure and add bulky substituents, Word keeps adjacent data columns aligned.
2. **Keep Paragraph Line Spacing on `Multiple` or `Single` (Prevent Top Clipping!):** If your table or body style uses `Exactly` (e.g., *Exactly 12 pt*) line spacing, Word slices off the top half of tall inline structures. Highlight the structure cells ➔ **Home ➔ Line and Paragraph Spacing ➔ Line Spacing Options** ➔ choose **Single** or **Multiple** (e.g., `1.15`), and click OK.
3. **Match Shaded Row Backgrounds:** If your SAR table has gray headers or alternating row shading, click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Standard Color (Transparent)]`** so structures sit cleanly over the cell fill, or choose **`[White Background (Color)]`** for crisp white framing.

::: tip Power Tip: Convert Pre-Typed Name or SMILES Columns Row by Row
If you already drafted a Word table with a `SMILES` or `Compound` column (`Aspirin`, `Caffeine`, `50-78-2`), highlight the text in a cell and click **`[SMILES]` ➔ `[Paste Structure]`** (or right-click ➔ **`Chemical Structure` ➔ `Paste Structure`**) to swap each cell into an aspect-ratio-locked structure in `0.1s`.
:::

## 2. Zero-Typo Experimental & Table Data via `Molecule Info`

Select any inserted structure picture—**or simply highlight a SMILES string or compound name in your manuscript prose**—and click **`[SMILES]` ➔ `[Molecule Info]`** (or right-click ➔ **`Chemical Structure` ➔ `Copy Molecule Info`**). All 12 property columns are computed locally on your PC:

<FeatureMedia id="common-molecule-info" />

| Manuscript Section You Are Writing | Best `Molecule Info` Option | What Gets Pasted (<kbd>Ctrl</kbd> + <kbd>V</kbd>) |
| :--- | :--- | :--- |
| **Starting a New 12-Column Property Table** | **`Copy All Summary (Text)`** | A formatted 2-row Word table (12 headers + values) copied via rich HTML + TSV |
| **Appending Rows to an Existing Word Table** | **`Copy as Excel Row (TSV Columns)`** | 12 tab-separated values without headers—select empty cells across a row and paste |
| **Experimental Section & HRMS (`m/z`)** | **`Copy Chemical Formula`**, **`Copy Molecular Weight (MW)`**, or **`Detailed Molecule Info Window...`** | Exact Hill formula (`C9H8O4`), MW (`180.16`), and 4-decimal Monoisotopic Exact Mass (`180.0423 Da`) for `[M+H]+` checks |
| **Supporting Information (SI) Compound Block** | **`Copy as Multi-Line List`** | Clean line-by-line list including Name, IUPAC Name, descriptors, SMILES, InChI, and InChIKey |

::: tip Power Tip: Extract Properties or SVG Directly From Plain Text
You do not even need to insert a structure picture first. Highlight any SMILES string or known compound name (e.g., `Ibuprofen`) in Word text, then run **`[Molecule Info]`** or **`[Copy SVG]`**—`func_SMILES2IMG` resolves the structure and copies the exact properties or vector SVG straight to your clipboard.
:::

## 3. Importing SAR Data Straight From Excel

Already calculated your library's MW, LogP, TPSA, `Fsp3`, and substructure match flags in Excel?

* **Paste Formatted Table Rows:** Select your calculated property rows in Excel, copy (<kbd>Ctrl</kbd> + <kbd>C</kbd>), paste into Word (**Keep Source Formatting** or **Match Destination Table Style**), lock **Fixed Column Width**, then populate the structure column with **`[Insert / Edit Structure]`** or **`[Paste Structure]`**.
* **Re-check Individual Compounds in Word:** Whenever you modify a structure in Word later by double-clicking, click **`[Molecule Info]`** right away to update its Formula, MW, and Exact Mass in the neighboring cell without returning to Excel.

## Related Guides

- [**Style & Export Reference**](/editor/publishing) — Exact clipboard output samples for all 5 `Molecule Info` formats (Summary, TSV, Multi-Line, Markdown, JSON)
- [**4. Journal Finish, SVG, Sharing**](/word/publishing) — Restyle all table and prose structures to ACS monochrome B&W in 1 click before submission
- [**Table Clipping & Batch Style Troubleshooting**](/word/troubleshooting/table-layout) — Fixes for clipped table images, shading contrast, and `Batch Style` selection scope
