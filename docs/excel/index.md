# Excel Overview

Use **`func_SMILES2IMG`** in Microsoft Excel to render 2D chemical structures inside cells, compute molecular properties and Lipinski drug-likeness across hundreds of rows locally, screen libraries with SMARTS, and import/export `.sdf` chemical databases.

<FeatureMedia id="excel-overview" />

::: tip Download the 10-Tab Self-Contained Excel Playbook
Explore all 40 formulas, real-world lab templates, and offline reference tables directly inside Excel—no web docs needed:  
<a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx</strong></a> · [See 10-tab breakdown & previews →](/excel/playbook)
:::

## 3 Ways to Work with Structures in Excel

| Method | How It Works | Best For |
| :--- | :--- | :--- |
| **Cell Formulas** | `=SMILES2IMG(A2)`, `=NAME2IMG("Aspirin")`, `=KET2IMG(A2)` | Batch rendering hundreds of rows that sort and filter cleanly |
| **Top `[SMILES]` Ribbon** | **`[Insert / Edit (SMILES)]`** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>) or **`[Paste Structure]`** | Sketching molecules/reactions or pasting clipboard structures |
| **Mouse Actions** | **Double-click** any structure cell to edit; **Right-click** ➔ `Chemical Structure` | Fast in-place edits, SVG copy, or opening Molecule Info |

::: info Microsoft 365 / Excel 2024+ vs. Excel 2016–2021 (`.FLOAT`)
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)` for native **in-cell images** that sort and filter with table rows.
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** (`=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)`) to create cell-anchored floating shapes (`xlMoveAndSize`). See [Legacy Excel (.FLOAT)](/excel/legacy-float).
:::

## Where to Go Next

- [**Structure Tables & Name/CAS**](/excel/structures) — Render columns from SMILES, compound names, solvents, and CAS numbers
- [**Styles, Highlighting & Reactions**](/excel/styles-reactions) — ACS monochrome, Dark Mode, SMARTS highlights, NMR atom numbers, and `=KET2IMG`
- [**Properties & Lipinski Ro5**](/excel/properties) — Molecular weight, exact mass, LogP, TPSA, rings, Fsp3, and chiral centers
- [**40 Formulas Reference**](/excel/formulas) — Complete parameter and return-type cheat sheet
