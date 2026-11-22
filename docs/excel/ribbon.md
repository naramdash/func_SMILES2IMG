# Ribbon, Context Menus, Diagnostics

Complete command reference for the Excel **`[SMILES]`** ribbon tab, the 3-level **`Chemical Structure`** right-click context menus, and the **`Diagnostics`** parser log window.

<FeatureMedia id="excel-ribbon-tools" />

## 1. Top `[SMILES]` Ribbon Commands

| Ribbon Group | Button / Menu | What It Does in Excel |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit (SMILES)` | Opens Ketcher 2D and writes `=SMILES2IMG("...")` (or `=KET2IMG` for reactions) into the active cell |
| | `Insert / Edit (InChI)` | Opens Ketcher 2D and writes `=INCHI2IMG("...")` into the active cell |
| | `Insert / Edit (KET)` | Opens Ketcher 2D and writes lossless `=KET2IMG("...")` JSON into the active cell |
| | `3D Structure (Mol*)` | Opens Mol\* 3D (auto-loads any 4-character PDB ID in the active cell) to insert a `.glb` 3D model or `.png` picture |
| | `Paste Structure` | Instantly converts clipboard or selected-cell SMILES/InChI/KET text into a rendered structure formula in 0.1s |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` | Extracts canonical SMILES, IUPAC InChI, or KET JSON from the selected formula cell or picture |
| | `Copy SVG` | Copies lossless vector SVG graphics + XML text to the clipboard for Illustrator, Inkscape, or Office |
| | `Molecule Info` | Copies individual properties, 17-column TSV/Markdown/Multi-Line summaries, or opens the 17-property dialog |
| | `Batch Style` | Converts selected cells (or the entire sheet) to `Standard Color`, `Publication B&W`, `White Background`, or `Dark Mode` |
| | `SDF File` | **Import SDF File (`.sdf`)** into a 12-column live sheet or **Export Sheet to SDF (`.sdf`)** (V2000) |
| **Diagnostics** | `Diagnostics` | Opens the EPAM Indigo parser log window to inspect exact chemical syntax errors behind `#VALUE!` cells |
| **About** | `Online Docs` | Opens the official Excel documentation portal (`/excel/`) in your web browser |
| | `About func_SMILES2IMG` | Displays installed version, embedded engine versions, open-source licenses, and 1-click update check |

## 2. Right-Click Context Menus (3 Targets)

Right-clicking a **worksheet cell**, an **Excel Table (`ListRange`)**, or a **floating `.FLOAT` / 3D picture** exposes the 3-tier **`Chemical Structure`** context menu:

![Excel 3-tier Chemical Structure context menu](/excel-context-menu-guide.png)

## 3. Inspecting `#VALUE!` Syntax Errors with `Diagnostics`

When a cell formula returns `#VALUE!`, click **`[SMILES]` ➔ `[Diagnostics]`** on the ribbon. The log window records the exact cell address, input string, and EPAM Indigo parser message (for example, `unclosed ring #1` or `unmatched parenthesis`), allowing you to pinpoint malformed SMILES or SMARTS strings immediately.

## Related Guides

- [**Batch Style, TSV, 3D**](/excel/tools-export) — Practical workflows for sheet-wide styling, vector export, and 3D PDB models
- [**SDF Import & Export**](/excel/sdf) — 12-column `.sdf` worksheet import and V2000 `.sdf` export
- [**Formula Errors**](/excel/troubleshooting/formula-errors) — Troubleshooting guide for `#NAME?`, `#VALUE!`, and `#N/A` codes
