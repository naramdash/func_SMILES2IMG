# Excel Playbook (`.xlsx`)

Practice all 40 Excel chemical formulas, medicinal chemistry screening tables, SMARTS scaffold highlighting, KET reaction schemes, and offline reference sheets directly inside Microsoft Excel—no web documentation required.

<PlaybookPreview app="excel" />

::: tip How to Use This Playbook
1. Make sure **`func_SMILES2IMG`** is installed ([Installation](/guide/install)).
2. Download <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx (.xlsx)</strong></a> and open it directly in **Microsoft Excel**.
3. Follow one of the recommended 3-minute practice tracks below, or copy the ready-made stoichiometry and SAR templates straight into your own lab workbooks.
:::

## Recommended Practice Tracks

| Track | Tabs to Open | What You Will Master in 3 Minutes |
| :--- | :--- | :--- |
| **1. First-Time Quickstart** | `0. Start Here` ➔ `1. Structure Rendering` | Edit a SMILES/Name cell, resize row height, and **double-click** a structure cell to edit in **Ketcher 2D** |
| **2. Organic Synthesis Lab (ELN)** | `6. Reactions & KET` ➔ `7. Real-World Workflows` | Double-click a multi-step `=KET2IMG` synthesis scheme and auto-calculate reagent masses (`mg`) and `% yield` from `=SMILES2MW` |
| **3. Medicinal Chemistry & SAR** | `2. Properties & Lipinski` ➔ `3. SAR & Similarity` | Change the query SMARTS cell (`$B$3`) to watch `=SMILES.MATCH`, red substructure highlights (`hl:`), and Tanimoto similarity update live |

## Sheet-by-Sheet Breakdown (10 Tabs)

| Sheet Tab | What You Can Practice & Copy | Key Formulas / Features |
| :--- | :--- | :--- |
| **`0. Start Here`** | 3-minute interactive sandbox & 10-tab learning roadmap | Top 10 live formulas & double-click editing |
| **`1. Structure Rendering`** | Render 2D structures from common names, CAS numbers, SMILES, and IUPAC InChI | `=NAME2IMG`, `=SMILES2IMG`, `=INCHI2IMG`, `=NAME2SMILES` |
| **`2. Properties & Lipinski`** | 18-column medicinal chemistry screening table + live `AVERAGE` / `COUNTIF` KPI row | `=SMILES2FORMULA`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2HEAVY`, `=SMILES2FSP3`, `=SMILES2LIPINSKI` |
| **`3. SAR & Similarity`** | Editable query SMARTS (`$B$3`), red scaffold highlighting, Tanimoto rank & 12-group SMARTS dictionary | `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2IMG(..., "hl:...")` |
| **`4. Notation Conversions`** | Cross-notation standardization & automatic duplicate detection via `InChIKey` | `=SMILES2INCHI`, `=SMILES2INCHIKEY`, `=INCHI2SMILES`, `=INCHI2MW` |
| **`5. Styles & Transforms`** | 16 presets: ACS B&W, Dark Mode, custom ink, NMR numbers, stereo, flip & rotation | `"bw"`, `"white"`, `"ink=#1E40AF"`, `"num"`, `"stereo"`, `"flip rot90"` |
| **`6. Reactions & KET`** | Real Aspirin & Paracetamol synthesis reaction schemes and single-molecule KET canvases | `=KET2IMG`, `=KET2FORMULA`, `=KET2MW`, double-click Ketcher editing |
| **`7. Real-World Workflows`** | Copy-ready Organic Stoichiometry Scale-Up calculator (`mg` & `% yield`) + Hit-to-Lead MPO Triage | Live `=SMILES2MW(C6)*H6` scaling & `IF(AND(...))` candidate scoring |
| **`8. Legacy Excel (.FLOAT)`** | Cell-anchored floating picture shapes for Excel 2016, 2019, and 2021 compatibility | `=SMILES2IMG.FLOAT`, `=NAME2IMG.FLOAT` |
| **`9. Ribbon, SDF & Ref`** | Complete `[SMILES]` Ribbon & 3-level context menu map, `.sdf` 12-col guide, all 40 UDFs & error fixes | Full offline desktop reference (`#N/A`, `#VALUE!`, `#NAME?`) |

## Related Guides

- [**All 40 Functions**](/excel/formulas) — Full parameter reference, data types, and shared styling tokens
- [**SMILES, InChI**](/excel/structures) · [**Names, Solvents, CAS**](/excel/name-cas) — Step-by-step setup for structure tables and zero-SMILES name lookups
- [**Interactive Playbooks Hub**](/guide/playbooks) — Compare all three Office Playbooks ([Word .docx](/word/playbook) · [PowerPoint .pptx](/powerpoint/playbook))
