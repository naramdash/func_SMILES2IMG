# Formula Errors (`#NAME?`, `#VALUE!`, `#N/A`)

Diagnosis and fixes for every error or status code returned by `func_SMILES2IMG` formulas in Microsoft Excel.

## 1. `#NAME?` Error

Excel displays `#NAME?` when it does not recognize the function name:

| Cause | How to Tell | Fix |
| :--- | :--- | :--- |
| **Excel 2016 / 2019 / 2021** | Native in-cell image formulas (`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`) require Microsoft 365 or Excel 2024+. | Append **`.FLOAT`** to the formula: `=SMILES2IMG.FLOAT(A2)` or `=NAME2IMG.FLOAT("Aspirin")`. |
| **Add-In Not Loaded or Blocked** | Even `=SMILES2MW("CCO")` returns `#NAME?`, or the `[SMILES]` tab is missing. | Check [Installation Blocks](/guide/troubleshooting/install) and [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab). |

## 2. `#VALUE!` Error

Excel displays `#VALUE!` when the formula ran, but the input or query string is invalid:

1. **Empty or Over-Limit Input:** Blank/whitespace-only cells and inputs exceeding **16,384 characters** return `#VALUE!`.
2. **Check Chemical Syntax & Valence:** Look for unclosed ring numbers (e.g., `c1ccccc` missing the closing `1`), unmatched or empty parentheses (`CC(=O`), bad atomic valence (e.g., neutral pentavalent carbon `C(C)(C)(C)(C)C`), non-InChI strings passed to `INCHI2*`, non-KET JSON passed to `KET2*`, or invalid SMARTS patterns in `*.MATCH`.
3. **Empty or Merged Cells:** When referencing a merged range, always point to its **top-left cell**.
4. **Inspect Exact Parser Logs (`Diagnostics`):** Click **`[SMILES]` ➔ `[Diagnostics]`** on the ribbon to read the failing formula function name and exact EPAM Indigo parser error message (for example, `SMILES2IMG: Indigo: ...`).

## 3. `#N/A` Status & Guardrails

* **Brief Async Indicator:** When rendering a structure for the first time—or looking up an uncached compound name/CAS from PubChem—Excel briefly shows `#N/A` and **automatically replaces it** once complete.
* **Unknown Names & Invalid CAS Numbers:** If a compound name is not in the built-in offline dictionary (or PubChem is unreachable/returns no match), or if a CAS number fails mathematical check-digit validation (e.g., `"50-78-3"`), `=NAME2*` immediately returns `#N/A` without waiting on a failed network call.
* **Queries & Reactions in Single-Molecule Descriptors:** Passing query atoms / R-groups (`*`) or reaction schemes (`>>`) to single-molecule scalar property functions (`=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2FSP3`, etc.) returns `#N/A` instead of miscomputing a false number.

## Related Guides

- [**Image Sizing & .FLOAT**](/excel/troubleshooting/image-sizing) — Resize small structure cells cleanly and handle Excel 2016–2021 `.FLOAT` shapes
- [**All 49 Functions**](/excel/formulas) — Verify exact formula signatures, return data types, and style keywords
