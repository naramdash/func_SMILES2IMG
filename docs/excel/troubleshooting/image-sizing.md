# Image Sizing & `.FLOAT` Mode

Solutions for small or cramped molecule images in Excel cells, Excel 2016–2021 `.FLOAT` shape behavior, and sharing `.xlsx` workbooks externally.

## 1. Molecule Image Looks Too Small

In-cell chemical images automatically scale to fit inside their host cell while locking their chemical aspect ratio so bond angles never stretch.

* **Fix:** Increase **row height** (recommended `62–100 pt`) and **column width**.
* **Batch Tip:** Select all compound rows together, drag a single row divider to your target height, and every structure scales sharply at once.

## 2. Excel 2016 / 2019 / 2021 (`.FLOAT` Shapes)

Older perpetual Excel versions do not support native in-cell image formulas (`=SMILES2IMG`).

* Use the `.FLOAT` variants instead: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT("Aspirin")`, `=KET2IMG.FLOAT(A2)`, `=INCHI2IMG.FLOAT(A2)`.
* Generated picture shapes are anchored to their host cell (`xlMoveAndSize`), so resizing the cell scales the structure automatically.

## 3. Sharing `.xlsx` Files Without `#NAME?`

* Cached cell images and calculated descriptor values remain visible on other PCs until someone triggers a recalculation (<kbd>F9</kbd>).
* Before sending a final workbook to colleagues without `func_SMILES2IMG`, copy calculated numeric/text columns (<kbd>Ctrl</kbd> + <kbd>C</kbd>) ➔ **Paste as Values**, or export to **PDF**. See [Document Sharing](/guide/troubleshooting/document-sharing) for full details.

## Related Guides

- [**Formula Errors**](/excel/troubleshooting/formula-errors) — Resolve `#NAME?`, `#VALUE!`, `#N/A`, and `Diagnostics` parser errors
- [**Legacy Excel (.FLOAT)**](/excel/legacy-float) · [**SMILES, InChI**](/excel/structures) — Structure tables and cell-anchored floating shape guides
