# Substructure Search (SAR & SMARTS)

Output native Excel **`TRUE` / `FALSE` filter flags** across all three chemical notations (**`=SMILES.MATCH`**, **`=INCHI.MATCH`**, **`=KET.MATCH`**), and visually highlight matching atoms and bonds inside rendered structures (`hl:`).

<FormulaTable
  title="SAR Substructure & Highlight Sheet"
  :columns='["Target", "Query", "Formula"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1", "=SMILES.MATCH(A2, $B$2)"], desc: "Boolean SMARTS substructure match (benzene)", out: "TRUE" },
    { cells: ["InChI=1S/C9H8O4/c1-6(10)13-8-5-3-2-4-7(8)9(11)12/h2-5H,1H3,(H,11,12)", "", "=INCHI.MATCH(A3, $B$2)"], desc: "Direct InChI substructure match", out: "TRUE" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES2IMG(A4, \"\", \"hl:\"&$B$2)"], desc: "Visually highlight matching substructure atoms", out: "Highlighted 2D" }
  ]'
/>

<FeatureMedia id="excel-sar-screening" />

## 1. Dynamic Query Cell Pattern (Interactive SAR Dashboard)

Instead of hardcoding a SMARTS pattern in every row, place your query motif in a single header cell (e.g., `$B$3`) and reference it across your table:

| Column Purpose | Formula Referencing `$B$3` | What Happens When You Edit `$B$3` |
| :--- | :--- | :--- |
| **SMILES Substructure Match (`TRUE`/`FALSE`)** | `=SMILES.MATCH(A6, $B$3)` | Instantly flags every SMILES row containing the new query motif |
| **InChI / KET Parity Match (`TRUE`/`FALSE`)** | `=INCHI.MATCH(A6, $B$3)` · `=KET.MATCH(A6, $B$3)` | Screens native InChI or KET canvas columns directly |
| **Highlighted Structure Image** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | Colors the matching substructure atoms and bonds inside every cell image |

Combine `=SMILES.MATCH` with exact physicochemical thresholds to prioritize drug-like hits containing your required pharmacophore:

```text
=IF(AND(SMILES2MW(A2)<=500, SMILES2LOGP(A2)<=4, SMILES.MATCH(A2, "c1ccncc1")), "Priority Lead", "Exclude")
```

::: info SMARTS Matching Convention
Standard Daylight/RDKit **SMARTS** is parsed first. Uppercase `C` / `N` match aliphatic atoms and lowercase `c` / `n` match aromatic atoms; use atomic numbers (`[#6]`, `[#7]`) when matching either. Target aromaticity is normalized under EPAM Indigo's `generic` model so Kekule and lowercase aromatic SMILES match identically.
:::

## 2. Copy-Ready Dictionary: 12 Common Functional Group SMARTS

Use these standard SMILES/SMARTS patterns with `=SMILES.MATCH` (`INCHI.MATCH` / `KET.MATCH`) or `"hl:<pattern>"` across your spreadsheets:

| Functional Group / Ring | SMARTS Pattern | Substructure Match (`TRUE` / `FALSE`) | Highlighted Structure Formula |
| :--- | :--- | :--- | :--- |
| **Benzene Ring (Aromatic)** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **Pyridine Ring** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **Carboxylic Acid (`-COOH`)** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **Ester (`-COO-`)** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **Amide (`-CONH-`)** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **Sulfonamide (`-SO₂NH-`)** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **Ketone (`-C(=O)-`)** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **Aldehyde (`-CHO`)** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **Phenol (Aromatic `-OH`)** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **Nitro Group (`-NO₂`)** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **Any Halogen (`F, Cl, Br, I`)** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **Trifluoromethyl (`-CF₃`)** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

## Related Guides

- [**Identifier Conversion**](/excel/data-sdf) — Standardize Name/CAS/SMILES/KET/InChIKey strings and deduplicate compound tables
- [**LogP, TPSA & 3D**](/excel/lipinski-adme) — Pair SAR scaffold filters with LogP, TPSA, and Fsp3 cutoffs
- [**SDF Import**](/excel/sdf) — Import `.sdf` screening libraries directly into an 11-column live Excel SAR table
- [**Styles & Highlights**](/excel/styles-reactions) — Customize highlight options alongside `"bw"`, `"num"`, and dark mode
