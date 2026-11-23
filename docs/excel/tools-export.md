# Batch Style, 12-Column TSV, 3D (`Mol*`)

Losslessly restyle formula-backed chemical tables across an entire Excel worksheet or selection with Batch Style, expand a molecule cell into a 12-column TSV descriptor row, and place interactive 3D protein models (`Mol*`) from PDB codes in cells.

## 1. Lossless Formula-Rewriting `Batch Style` Across Ranges or Sheets

Click **`[SMILES]` ➔ `[Batch Style]`** on the Excel ribbon to standardize the background and bond palette of your chemical structures in one step:

* **Lossless Formula Rewriting:** Rewrites standalone `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG` (and `.FLOAT`) formulas by keeping the first argument (including nested calls or quoted KET strings), updating `[bg]` and color/ink tokens in `[style]`, and **preserving structural tokens (`hl:`, `num`, `stereo`, `h`, `all-c`) and 4th-arg rotations (`rot90`, `rot180`, `rot270`)**.
* **Dynamic Highlight Compatibility & Composite Protection:** Even formulas with dynamic expressions like `=SMILES2IMG(C5, , "hl:" & $B$3)` keep `$B$3` live after Batch Style, while composite wrapper formulas (such as `=IF(..., SMILES2IMG(...))`) are safely skipped without corruption.
* **Scope:** Select a multi-cell range to restyle only that selection, select a single cell to scan all formula cells on the **entire active worksheet**, or click an inserted 2D shape to restyle only that shape.

| `Batch Style` Preset | Formula Written in Excel | Typical Spreadsheet Use Case |
| :--- | :--- | :--- |
| **`Standard Color (Transparent)`** | `=SMILES2IMG(A2, "trans", "color ...")` | Default IUPAC element colors over clean worksheet cells |
| **`Publication B&W (ACS Monochrome)`** | `=SMILES2IMG(A2, "trans", "bw ...")` | Journal submission tables, patent filings, and B&W printouts |
| **`White Background (Color)`** | `=SMILES2IMG(A2, "white", "color ...")` | Crisp opaque card over shaded or conditionally formatted rows |
| **`Dark Mode (White Ink)`** | `=SMILES2IMG(A2, "trans", "white ...")` | High-contrast white bond lines over dark dashboard tables |

## 2. 12-Column TSV Row Expansion & Formula Identifier Extraction

Select any formula structure cell (`=NAME2IMG`, `=SMILES2IMG`, `=INCHI2IMG`, `=KET2IMG`) or inserted picture shape and use **`[SMILES]`** (or right-click ➔ **`Chemical Structure`**) to pull underlying data back into Excel columns or graphics tools:

<FeatureMedia id="common-molecule-info" />

* **`Copy as Excel Row (TSV Columns)` (`Molecule Info`):** Pastes all **12 calculated descriptors** (`Name`, `Formula`, `MW`, `Exact Mass`, `LogP`, `TPSA`, `Heavy Atoms`, `Rings`, `Aromatic Rings`, `Fsp3`, `Canonical SMILES`, `InChIKey`) across adjacent columns in a single <kbd>Ctrl</kbd> + <kbd>V</kbd>—ideal when you want static values without writing 12 separate UDF columns ([full 5-format specification →](/editor/publishing)).
* **`Copy SMILES` / `Copy InChI` / `Copy KET`:** Extracts resolved canonical identifiers even from `=NAME2IMG("Aspirin")`, CAS lookup cells, or legacy `.FLOAT` shapes.
* **`Copy SVG`:** Copies vector graphics (`image/svg+xml` and raw `<svg>` XML text) so a highlighted hit from your screening table can be pasted directly into Adobe Illustrator or Inkscape without pixelation.

## 3. Cell-Driven 3D Proteins via `Mol*`

Embed 360°-rotatable 3D macromolecular models (`.glb`) or transparent 2D protein figures (`.png`) directly over an Excel sheet using the built-in **Mol\*** viewer:

<FeatureMedia id="common-molstar-3d" />

1. **Auto-Detect PDB ID from Cell:** Type a 4-character PDB code (such as `1CRN`, `4HHB`, `6LU7`, or `1BNA`) in a cell, keep that cell selected, and click **`[SMILES]` ➔ `[3D Structure (Mol*)]`**—Mol\* reads the cell value and fetches that PDB entry immediately. *(From a blank cell, you can search any PDB/AlphaFold code or open a local `.pdb` / `.cif` / `.sdf` file offline.)*
2. **Choose Mode & Representation:** Configure **Cartoon**, **Surface**, or **Ball & Stick** views and select **`3D Model (GLB)`** for live rotation on the worksheet or **`2D Picture`** for a lightweight transparent PNG.
3. **Edit Later at the Same Angle:** Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Worksheet`**). **Double-click** the placed protein object anytime to reopen Mol\* with your saved representation and camera viewpoint.

## Related Guides

- [**Ribbon & Context Menus**](/excel/ribbon) — Complete Excel `[SMILES]` ribbon map, 3 cell/picture context menus, and the Indigo parser log
- [**SDF Import**](/excel/sdf) — Import `.sdf`/`.mol` files into an 11-column live worksheet
- [**Mol\* 3D Viewer Manual**](/editor/molstar) — Protein representations, ligand binding sites, and offline local PDB/CIF loading
