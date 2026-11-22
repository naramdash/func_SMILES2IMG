# Reaction Schemes (KET)

Embed complete, multi-step organic synthesis pathways—including reaction arrows (`➔`), `+` signs, catalysts, solvents, and stereochemical layouts—directly inside Excel cells via `=KET2IMG` while keeping 100% lossless vector/coordinate editing in **Ketcher 2D**.

<FormulaTable
  title="Reaction Scheme (KET) Sheet"
  :columns='["Reaction / KET", "Formula"]'
  :rows='[
    { input: "CC(=O)O.OCC>>CC(=O)OCC", formula: "=KET2IMG(A2)", desc: "Render reaction scheme in standard IUPAC color", out: "Color reaction" },
    { input: "CC(=O)O.OCC>>CC(=O)OCC", formula: "=KET2IMG(A3, \"\", \"bw\")", desc: "Render monochrome ACS reaction scheme", out: "B&W reaction" }
  ]'
/>

<FeatureMedia id="excel-reactions-ket" />

## 1. How to Insert & Re-Edit a Reaction Scheme Cell

1. Select a blank target cell and click **`[SMILES]` ➔ `[Insert / Edit (KET)]`** on the Excel ribbon (note: **`[Insert / Edit (SMILES)]`** also automatically saves as `=KET2IMG(...)` whenever reaction arrows or multiple components are drawn).
2. Draw reactants, reaction arrows (`➔`), `+` symbols, catalysts/temperatures above/below the arrow, and products inside the offline **Ketcher 2D** editor.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to embed the scheme as `=KET2IMG("...")` (or `=KET2IMG.FLOAT("...")` on Excel 2016–2021).
4. **Double-click** that cell anytime to reopen Ketcher in `0.05s`—every arrow angle, reagent layout, stereocenter, and text annotation reloads intact for instant updates.

::: tip Auto-Detection with `=SMILES2IMG`
Even if a column mixes single-molecule SMILES strings and Ketcher KET strings, `=SMILES2IMG(A2)` automatically detects KET JSON and routes it through the Ketcher reaction renderer seamlessly.
:::

## 2. Applying Shared Styles (`[bg]`, `[style]`) to Reaction Schemes

Because `=KET2IMG` and `=KET2IMG.FLOAT` use the exact same styling parser as `=SMILES2IMG`, you can adapt an entire synthesis route for manuscripts or dark slides without redrawing it:

| Use Case | Excel Formula | Result |
| :--- | :--- | :--- |
| **Standard IUPAC Color (Transparent)** | `=KET2IMG(A2)` | Colored heteroatoms (`O` red, `N` blue) with transparent background |
| **ACS Journal Monochrome B&W** | `=KET2IMG(A2, "", "bw")` *(or `=KET2IMG(A2, "bw")`)* | High-contrast black-and-white publication reaction scheme |
| **Dark Slide / Dashboard White Ink** | `=KET2IMG(A2, "#0F172A", "white")` | Crisp white arrows, bonds, and conditions over dark backgrounds |
| **Legacy Excel (2016–2021)** | `=KET2IMG.FLOAT(A2, "", "bw")` | Cell-anchored floating reaction image (`xlMoveAndSize`) |

## 3. Extracting Data from Single-Molecule KET Cells

When a cell contains a single molecule stored as KET JSON (e.g., drawn with custom atom positions or explicit stereowedges in Ketcher), you can also query its descriptors directly:

* **Formula / MW / Exact Mass:** `=KET2FORMULA(A2)`, `=KET2MW(A2)`, `=KET2EXACTMASS(A2)`
* **Identifiers & Name:** `=KET2NAME(A2)`, `=KET2INCHI(A2)`, `=KET2INCHIKEY(A2)`
* **ADME & Lipinski Rules:** `=SMILES2LIPINSKI(A2)`, `=SMILES2LOGP(A2)`, `=SMILES2TPSA(A2)` *(all `SMILES2*` descriptor functions accept single-molecule KET cells automatically)*

## Related Guides

- [**Styles & Highlights**](/excel/styles-reactions) — Complete list of shared `[bg]`, `[style]`, and `[transform]` keywords across all 8 image formulas
- [**SMILES, InChI**](/excel/structures) — Multi-row SMILES/InChI tables and `0.05s` double-click editing
- [**Ketcher 2D Sketcher Manual**](/editor/) — Toolbar shortcuts for reaction arrows, atom-mapping, and stereocenters
