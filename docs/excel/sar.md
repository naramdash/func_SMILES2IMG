# Substructure & Similarity (SAR)

Output native Excel **`TRUE` / `FALSE` filter flags (`=SMILES.MATCH`)** and **`0.0–1.0` 2D similarity scores (`=SMILES.SIMILARITY`)** across thousands of molecules (SMILES, InChI, or KET), and visually highlight matching atoms inside `...2IMG` structures (`hl:`).

<FormulaTable
  title="SAR Substructure, Similarity, Highlight Sheet"
  :columns='["Target", "Reference", "Formula"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES.MATCH(A2, \"c1ccccc1\")"], desc: "Boolean substructure match (benzene)", out: "TRUE" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1C(=O)O", "=SMILES.SIMILARITY(A3, $B$3)"], desc: "2D Tanimoto similarity vs Reference ($B$3)", out: "0.00–1.00" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES2IMG(A4, \"\", \"hl:c1ccccc1\")"], desc: "Visually highlight matching benzene atoms", out: "Highlighted 2D" }
  ]'
/>

<FeatureMedia id="excel-sar-screening" />

## 1. Dynamic Query Cell Pattern (Interactive SAR Dashboard)

Instead of hardcoding a SMARTS string in every row, place your query scaffold in a single header cell (e.g., `$B$3`) and reference it across your table:

| Column Purpose | Formula Referencing `$B$3` | What Happens When You Edit `$B$3` |
| :--- | :--- | :--- |
| **Scaffold Match (`TRUE`/`FALSE`)** | `=SMILES.MATCH(A6, $B$3)` | Instantly flags every row containing the new query motif |
| **Highlighted Structure Image** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | Colors the matching substructure atoms and bonds inside every cell image |
| **2D Similarity Score (`0.0–1.0`)** | `=SMILES.SIMILARITY(A6, $B$3, "tanimoto")` | Ranks every compound by 2D fingerprint similarity (`"tanimoto"`, `"tversky"`, or `"euclid-sub"`) |

Combine `=SMILES.MATCH` with `=SMILES2LIPINSKI` to prioritize drug-like hits containing your required pharmacophore:

```text
=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "Priority Lead", "Exclude")
```

## 2. Copy-Ready Dictionary: 12 Common Functional Group SMARTS

Use these standard SMILES/SMARTS patterns with `=SMILES.MATCH` or `"hl:<pattern>"` across your spreadsheets:

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

- [**Identifier Conversion**](/excel/data-sdf) — Standardize Name/CAS/SMILES/InChIKey strings and deduplicate compound tables
- [**Lipinski, ADME, 3D**](/excel/lipinski-adme) — Pair SAR scaffold filters with Lipinski Ro5, TPSA, and Fsp3 cutoffs
- [**SDF Import & Export**](/excel/sdf) — Import `.sdf` screening libraries directly into a 12-column live Excel SAR table
- [**Styles & Highlights**](/excel/styles-reactions) — Customize highlight options alongside `"bw"`, `"num"`, and dark mode
