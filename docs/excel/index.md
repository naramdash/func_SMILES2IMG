# Excel Overview

Use **`func_SMILES2IMG`** in Microsoft Excel to render crisp 2D chemical structures inside cells, compute molecular properties and 3D complexity locally across rows, screen compound libraries with SMARTS, and import `.sdf` chemical databases.

<FeatureMedia id="excel-overview" />

::: tip Download the 10-Tab Self-Contained Excel Playbook
Practice all 49 formulas, organic stoichiometry calculators, SAR dashboards, and offline reference tables directly inside Excel—no separate web documentation required:  
<a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx</strong></a> · [See 10-tab breakdown & recommended tracks →](/excel/playbook)
:::

## Quick Start

Click **Copy All** and paste into cell **`A1`** in Excel (or click any row to copy its input + formula; on **Excel 2016–2021**, append `.FLOAT`—e.g., `=SMILES2IMG.FLOAT(A2)`):

<FormulaTable
  title="Quick Start Sheet"
  :columns='["SMILES / Input", "Formula"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG(A2)", desc: "Render 2D structure from SMILES in A2", out: "2D Ethanol" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A3)", desc: "Calculate molecular weight locally", out: "180.16" },
    { input: "", formula: "=NAME2IMG(\"Aspirin\")", desc: "Standalone 2D structure from Name or CAS #", out: "2D Aspirin" }
  ]'
/>

Increase **row height** (`62–80 pt`) and **column width** to scale the molecule cleanly with locked aspect ratio. **Double-click** any rendered cell to reopen the offline **Ketcher 2D** editor and edit in place.

## 3 Ways to Work in Excel

| Method | How It Works | Best For |
| :--- | :--- | :--- |
| **Cell Formulas** | `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=SMILES2MW(B2)` | Automated tables across rows that sort and filter cleanly |
| **Top `[SMILES]` Ribbon** | **`[Insert / Edit]`**, **`[Paste Structure]`**, **`[3D Structure (Mol*)]`** | Sketching molecules/reactions, clipboard conversion, or 3D PDB models |
| **Mouse Actions** | **Double-click** to edit in-place; **Right-click** ➔ `Chemical Structure` | Direct structure edits, vector SVG copy, or opening `Molecule Info` |

::: info Microsoft 365 / Excel 2024+ vs. Excel 2016–2021 (`.FLOAT`)
* **Microsoft 365 / Excel 2024+:** `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)` create native **in-cell images** that sort and filter with table rows.
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** (`=SMILES2IMG.FLOAT(A2)`) to create cell-anchored floating shapes (`xlMoveAndSize`). All property and SAR formulas (`=SMILES2MW`, `=SMILES.MATCH`, etc.) work identically on every Excel version without `.FLOAT`. See [Legacy Excel (.FLOAT)](/excel/legacy-float).
:::

## Explore by Section

| Section | Pages | Key Formulas & Ribbon Tools |
| :--- | :--- | :--- |
| **Getting Started** | [**Playbook (.xlsx)**](/excel/playbook) · [**Use Cases**](/excel/use-cases) · [**All 49 Functions**](/excel/formulas) | Unified `=<INPUT>2<OUTPUT>` map, 6 lab scenarios & 10-tab workbook |
| **Structure Images** | [**SMILES, InChI**](/excel/structures) · [**Names, Solvents, CAS**](/excel/name-cas) · [**Reaction Schemes (KET)**](/excel/reactions) · [**Styles & Highlights**](/excel/styles-reactions) | `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG`, `.FLOAT` |
| **Properties & 3D** | [**Formula, MW, Exact Mass**](/excel/properties) · [**LogP, TPSA & 3D**](/excel/lipinski-adme) | `=SMILES2MW`, `=INCHI2LOGP`, `=KET2TPSA`, `=SMILES2FSP3` |
| **SAR & Identifiers** | [**Substructure Search**](/excel/sar) · [**Identifier Conversion**](/excel/data-sdf) | `=SMILES.MATCH`, `=INCHI.MATCH`, `=KET2SMILES`, `=SMILES2INCHIKEY` |
| **Sheet Tools & 3D** | [**SDF Import**](/excel/sdf) · [**Batch Style, TSV, 3D**](/excel/tools-export) · [**Ribbon & Context Menus**](/excel/ribbon) | `[Import SDF]`, `[Batch Style]`, `[Copy TSV]`, `[3D Structure (Mol*)]` |
