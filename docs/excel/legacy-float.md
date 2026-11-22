# Legacy Excel (`.FLOAT`)

Native in-cell picture formulas (`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`) require **Microsoft 365 or Excel 2024+**. On **Excel 2016, 2019, and 2021**, append **`.FLOAT`** to render cell-anchored floating picture shapes instead.

<FormulaTable
  title="Excel 2016–2021 (.FLOAT) Sheet"
  :columns='["SMILES / Input", "Formula"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG.FLOAT(A2)", desc: "Cell-anchored floating PNG from SMILES", out: "xlMoveAndSize" },
    { input: "", formula: "=NAME2IMG.FLOAT(\"Caffeine\")", desc: "Standalone floating structure from Name", out: "xlMoveAndSize" }
  ]'
/>

<FeatureMedia id="excel-legacy-float" />

## 1. `.FLOAT` Formula Reference

| Modern Excel (365 / 2024+) | Legacy Excel (2016 / 2019 / 2021) | Example |
| :--- | :--- | :--- |
| `=SMILES2IMG(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2)` |
| `=NAME2IMG(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT("Caffeine")` |
| `=KET2IMG(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2)` |
| `=INCHI2IMG(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2)` |

## 2. How `.FLOAT` Shapes Behave

* Shapes are anchored to their host cell (`xlMoveAndSize`), so resizing row height or column width scales the structure cleanly while keeping aspect ratio locked.
* All shared `[background]`, `[style]`, and `[transform]` keywords (`"bw"`, `"white"`, `"hl:..."`, `"num stereo"`, `"ink=#HEX"`) work identically to the in-cell formulas.
* All property, Lipinski, SAR, and conversion formulas (`=SMILES2MW`, `=SMILES2LIPINSKI`, `=SMILES.MATCH`, `=NAME2SMILES`, etc.) return standard Excel numbers/text and work on **every Excel version (2016–365)** without `.FLOAT`.

## Related Guides

- [**SMILES, InChI**](/excel/structures) — Multi-row structure tables and double-click Ketcher editing
- [**Styles & Highlights**](/excel/styles-reactions) — Shared styling keywords across all `...2IMG` and `.FLOAT` functions
- [**Image Sizing & .FLOAT**](/excel/troubleshooting/image-sizing) — Resizing cells and sharing `.xlsx` workbooks cleanly
