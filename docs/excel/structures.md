# SMILES & InChI Structure Images

Render high-resolution 2D molecular structures directly inside Excel cells from **SMILES strings (`=SMILES2IMG`)** or **IUPAC InChI strings (`=INCHI2IMG`)**—scaling sharply with row height and sorting cleanly alongside table rows.

<FormulaTable
  title="SMILES & InChI Rendering Sheet"
  :columns='["Compound", "SMILES / InChI", "Formula"]'
  :rows='[
    { cells: ["Ethanol", "CCO", "=SMILES2IMG(B2)"], desc: "Render in-cell 2D structure from SMILES (Col B)", out: "2D Ethanol" },
    { cells: ["Ethanol", "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3", "=INCHI2IMG(B3)"], desc: "Render in-cell 2D structure from IUPAC InChI (Col B)", out: "2D Ethanol" }
  ]'
/>

<FeatureMedia id="excel-formula-table" />

## 1. Unified Image Syntax (`SMILES` & `InChI`)

Both formulas share the exact same **4-argument signature `(input, [background], [style], [transform])`** and run 100% offline in 0ms:

| Input Column | Modern Excel (365 / 2024+) | Legacy Excel (2016–2021) | Notes |
| :--- | :--- | :--- | :--- |
| **SMILES String** | `=SMILES2IMG(B2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(...)` | Also auto-detects single-molecule/reaction KET JSON |
| **IUPAC InChI** (`InChI=1S/...`) | `=INCHI2IMG(C2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(...)` | Renders standard PubChem/ChEMBL/NIST InChI strings directly |

## 2. Build a Multi-Row Structure Table

1. Enter compound names in column `A` and SMILES strings in column `B` (e.g., Aspirin `CC(=O)Oc1ccccc1C(=O)O`).
2. Type `=SMILES2IMG(B2)` (or `=INCHI2IMG(B2)`) in `C2` and press <kbd>Enter</kbd> (on Excel 2016–2021, append **`.FLOAT`**).
3. Set **row height** (`62–80 pt` recommended) and widen column `C`—the structure scales while locking bond aspect ratio and stays locked to its row when filtering or sorting.
4. Drag the fill handle down to render hundreds of compounds at once.

## 3. Sketch, Convert in 0.1s & Double-Click to Edit

| Action | How to Do It | Result |
| :--- | :--- | :--- |
| **Draw a Molecule** | Select a cell ➔ **`[SMILES]` ➔ `[Insert / Edit (SMILES)]`** (or `InChI`) ➔ sketch in Ketcher ➔ <kbd>Ctrl</kbd> + <kbd>Enter</kbd> | Writes `=SMILES2IMG("...")` (or `=INCHI2IMG(...)`) directly into the cell |
| **0.1s Instant Convert** | Copy a SMILES/InChI string or select a text cell ➔ **`[SMILES]` ➔ `[Paste Structure]`** | Converts clipboard or cell text immediately without opening the editor |
| **In-Place Re-Editing** | **Double-click** any rendered structure cell or `.FLOAT` picture | Reopens **Ketcher 2D** in `0.05s` with the molecule preloaded |

## Related Guides

- [**Names, Solvents, CAS**](/excel/name-cas) — Render `Aspirin`, `THF`, `HATU`, or `"50-78-2"` directly with `=NAME2IMG`
- [**Reaction Schemes (KET)**](/excel/reactions) — Store multi-step reactions cleanly inside cells with `=KET2IMG`
- [**Styles & Highlights**](/excel/styles-reactions) — Customize `[bg]`, `[style]` (`"bw"`, `"hl:..."`, `"num stereo"`), and `[transform]`
- [**Legacy Excel (.FLOAT)**](/excel/legacy-float) — Cell-anchored shapes for Excel 2016, 2019, and 2021
