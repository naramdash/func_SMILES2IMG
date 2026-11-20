# Formula Errors (`#NAME?`, `#VALUE!`, `#N/A`)

Diagnosis and fixes for every error or status code returned by `func_SMILES2IMG` formulas in Microsoft Excel.

## 1. `#NAME?` Error

Excel displays `#NAME?` when it does not recognize the function name:

| Cause | How to Tell | Fix |
| :--- | :--- | :--- |
| **Excel 2016 / 2019 / 2021** | Native in-cell image formulas (`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`) require Microsoft 365 or Excel 2024+. | Append **`.FLOAT`** to the formula: `=SMILES2IMG.FLOAT(A2)` or `=NAME2IMG.FLOAT("Aspirin")`. |
| **Add-In Not Loaded or Blocked** | Even `=SMILES2MW("CCO")` returns `#NAME?`, or the `[SMILES]` tab is missing. | Check [Installation Blocks](/guide/troubleshooting/install) and [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab). |

## 2. `#VALUE!` Error

Excel displays `#VALUE!` when the formula ran, but the input could not be parsed as a valid chemical structure.

1. **Check Chemical Syntax:** Look for unclosed ring numbers (e.g., `c1ccccc` missing the closing `1`), unmatched parentheses, or stray spaces/line breaks.
2. **Empty or Merged Cells:** When referencing a merged range, always point to its **top-left cell**.
3. **Inspect Exact Parser Logs (`Diagnostics`):** Click **`[SMILES]` ➔ `[Diagnostics]`** on the ribbon to read the exact EPAM Indigo parser error message for the failing cell.

## 3. `#N/A` Shown Momentarily (~0.1s)

* **Normal Async Indicator:** When rendering a new structure for the first time—or looking up an uncached compound name/CAS from PubChem—Excel briefly shows `#N/A` and **automatically replaces it** once rendering finishes.
* **Stuck on `#N/A` for `=NAME2IMG` / `=NAME2SMILES`:** The name may not exist in the built-in offline dictionary or PubChem, or your PC is offline for an uncached lookup. Try a standard IUPAC name, CAS number (`"50-78-2"`), or direct SMILES string.

## Related Guides

- [**Image Sizing & .FLOAT Mode**](/excel/troubleshooting/image-sizing) — Resize small structure cells cleanly and handle Excel 2016–2021 .FLOAT shapes.
- [**40 Formula & Style Cheat Sheet**](/excel/formulas) — Verify exact formula signatures, return data types, and style keywords.
