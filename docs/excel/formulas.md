# All 40 Excel Functions

Every **`func_SMILES2IMG`** formula follows a unified **`=<INPUT>2<OUTPUT>`** design: pick a function prefix matching your **input** (`SMILES`, `NAME`, `INCHI`, or `KET`) and a suffix matching the **cell output** you need—whether a crisp 2D structure graphic, a quantitative property, a SAR screening flag, or a converted identifier string.

::: tip Two Shared Pipelines Across All 40 Functions
* **Structure Image Output (`8` functions):** `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, and `=KET2IMG` (plus their `.FLOAT` equivalents) share the identical 4-argument signature `(input, [bg], [style], [transform])` and the same EPAM Indigo renderer. Only the 1st argument's input notation varies.
* **Scalar Output (`32` functions):** Property, Lipinski/ADME, SAR, and identifier conversion functions return native Excel numbers, booleans (`TRUE`/`FALSE`), or strings—ready for `IF`, `FILTER`, `SORT`, `XLOOKUP`, and stoichiometry math (`MW × mmol`). All `SMILES2*` descriptor functions also auto-detect InChI and single-molecule KET cells.
:::

## ⭐ Top 5 Ready-to-Copy Formulas

Click **Copy All** and paste into cell **`A1`** to populate inputs (`A2:A5`) alongside live formulas (`B2:B6`), or copy any row individually (on Excel 2016–2021, append **`.FLOAT`** to `...2IMG` functions):

<FormulaTable
  title="Top 5 Ready-to-Copy Lab Sheet"
  :columns='["SMILES / Input", "Formula"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG(A2)", desc: "2D IUPAC color structure from SMILES in A2", out: "2D Ethanol" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A3, \"\", \"bw\")", desc: "ACS journal monochrome B&W structure from A3", out: "B&W Aspirin" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A4)", desc: "Molecular weight in g/mol (0ms offline)", out: "180.16" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LIPINSKI(A5)", desc: "Lipinski Rule of 5 drug-likeness", out: "PASS" },
    { input: "", formula: "=NAME2IMG(\"Aspirin\")", desc: "Standalone 2D structure from name or CAS #", out: "2D Aspirin" }
  ]'
/>

## 🗺️ Guide Map by Category

| Sidebar Category | Included Functions (`40` Total) | Pages |
| :--- | :--- | :--- |
| **Structure Images** | `SMILES2IMG`, `NAME2IMG`, `INCHI2IMG`, `KET2IMG` (+ 4 `.FLOAT`) | [SMILES, InChI](/excel/structures) · [Names, Solvents, CAS](/excel/name-cas) · [Reaction Schemes (KET)](/excel/reactions) · [Styles & Highlights](/excel/styles-reactions) |
| **Properties & ADME** | `*2FORMULA`, `*2MW`, `*2EXACTMASS`, `*2NAME` (×3) + 11 ADME/3D descriptors | [Formula, MW, Exact Mass](/excel/properties) · [Lipinski, ADME, 3D](/excel/lipinski-adme) |
| **SAR & Identifiers** | `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=NAME2SMILES`, `*2INCHI`, `INCHI2SMILES`, `*2INCHIKEY` | [Substructure & Similarity](/excel/sar) · [Identifier Conversion](/excel/data-sdf) |
| **Sheet Tools & 3D** | Multi-compound `.sdf` I/O, `[Batch Style]`, 17-col TSV & 3D `Mol*` | [SDF Import & Export](/excel/sdf) · [Batch Style, TSV, 3D](/excel/tools-export) · [Ribbon & Context Menus](/excel/ribbon) |

## 1. Structure Image Functions (`8`)

All 4 input formats share `(input, [bg], [style], [transform])` ([SMILES, InChI](/excel/structures) · [Names, Solvents, CAS](/excel/name-cas) · [Reaction Schemes (KET)](/excel/reactions)):

| Input Notation (`Arg 1`) | Microsoft 365 / Excel 2024+ | Excel 2016–2021 | Example |
| :--- | :--- | :--- | :--- |
| **SMILES** *(auto-detects KET)* | `=SMILES2IMG(cell, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **Name or CAS #** | `=NAME2IMG(name_or_cas, [bg], [style], [transform])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **IUPAC InChI** | `=INCHI2IMG(cell, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |
| **Reaction / Canvas (KET)** | `=KET2IMG(cell, [bg], [style], [transform])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |

::: tip Shorthand Tips
* **Minimal Call:** Only the 1st argument is required: `=SMILES2IMG(A2)`.
* **2-Arg B&W Shortcut:** Pass `"bw"` (or `FALSE`) in the 2nd argument to render transparent black-and-white art immediately: `=SMILES2IMG(A2, "bw")`.
* **Delimiters:** Separate multiple `[style]` tokens with spaces, commas, pipes (`|`), semicolons (`;`), or `+`.
:::

## 2. Formula, MW, Exact Mass, Name (`12`)

Except for online common-name lookups (`*2NAME`, cached locally after 1 query), all property calculations run **0ms purely offline** ([Formula, MW, Exact Mass](/excel/properties)):

| Output | SMILES Input | InChI / KET Input | Type | Aspirin (`A2`) |
| :--- | :--- | :--- | :---: | :--- |
| **Substance Name** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `String` | `"Aspirin"` |
| **Chemical Formula (Hill)** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)` · `=KET2FORMULA(A2)` | `String` | `"C9H8O4"` |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)` · `=KET2MW(A2)` | `Double` | `180.16` |
| **Monoisotopic Exact Mass** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)` · `=KET2EXACTMASS(A2)` | `Double` | `180.0423` |

## 3. Lipinski Ro5, ADME, 3D Descriptors (`11`)

Every descriptor function accepts SMILES, InChI, or single-molecule KET cells automatically ([Lipinski, ADME, 3D](/excel/lipinski-adme)):

| Descriptor | Formula | Type | Aspirin | MedChem Guideline |
| :--- | :--- | :---: | :---: | :--- |
| **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `String` | `"PASS"` | MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10 |
| **LogP (Crippen)** | `=SMILES2LOGP(A2)` | `Double` | `1.31` | ≤ 5.0 recommended |
| **TPSA (Å²)** | `=SMILES2TPSA(A2)` | `Double` | `63.60` | ≤ 140 Å² oral; ≤ 90 Å² CNS BBB |
| **H-Bond Donors** | `=SMILES2HBD(A2)` | `Integer` | `1` | ≤ 5 (`-OH`, `-NH`) |
| **H-Bond Acceptors** | `=SMILES2HBA(A2)` | `Integer` | `4` | ≤ 10 (`N`, `O` acceptors) |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `Integer` | `3` | ≤ 10 recommended |
| **Heavy Atoms** | `=SMILES2HEAVY(A2)` | `Integer` | `13` | Non-hydrogen atoms |
| **Total Rings (SSSR)** | `=SMILES2RINGS(A2)` | `Integer` | `1` | Smallest Set of Smallest Rings |
| **Aromatic Rings** | `=SMILES2AROMATIC(A2)` | `Integer` | `1` | ≤ 3 recommended |
| **Fraction Csp³ (Fsp³)** | `=SMILES2FSP3(A2)` | `Double` | `0.11` | ≥ 0.25 (`0.00–1.00`) |
| **Chiral Centers** | `=SMILES2CHIRAL(A2)` | `Integer` | `0` | Defined tetrahedral stereocenters |

## 4. SAR Screening & Similarity (`2`)

Both SAR functions accept SMILES, InChI, or KET inputs ([Substructure & Similarity](/excel/sar)):

| Output | Formula | Type | Result | How It Works |
| :--- | :--- | :---: | :---: | :--- |
| **Substructure Match** | `=SMILES.MATCH(target, "query")` | `Boolean` | `TRUE` / `FALSE` | Checks if `target` contains the SMILES/SMARTS pattern `query` |
| **2D Fingerprint Similarity** | `=SMILES.SIMILARITY(mol1, mol2, [metric])` | `Double` | `0.0–1.0` | Computes 2D similarity (`"tanimoto"`, `"tversky"`, or `"euclid-sub"`) |

## 5. Identifier Conversions (`7`)

Standardize Chemical Name, CAS #, Canonical SMILES, IUPAC InChI, and 27-character InChIKey ([Identifier Conversion](/excel/data-sdf)):

| Conversion | Formula | Type | Example Output |
| :--- | :--- | :---: | :--- |
| **Name or CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` · `=KET2INCHI(A2)` | `String` | `"InChI=1S/C9H8O4/..."` |
| **InChI ➔ Canonical SMILES** | `=INCHI2SMILES(A2)` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` · `=INCHI2INCHIKEY(A2)` · `=KET2INCHIKEY(A2)` | `String` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

## 6. Shared Background (`[bg]`) & Style (`[style]`) Tokens

All 8 `...2IMG` and `...2IMG.FLOAT` functions share these exact styling parameters ([Styles & Highlights](/excel/styles-reactions)):

<FeatureMedia id="excel-style-matrix" />

| Parameter | Token / Value | Accepted Aliases | Effect |
| :--- | :--- | :--- | :--- |
| **Background `[bg]`** | `""` / `"trans"` · `"white"` · `#RRGGBB` | `"transparent"`, CSS color names | Transparent canvas, opaque white box, or custom background fill |
| **Color / Monochrome** | `color` · `bw` | `cpk`, `mono`, `black`, `grayscale` | Standard IUPAC element coloring or ACS B&W line art |
| **Custom Ink** | `white` · `ink=#HEX` | `light`, `line=#HEX`, `base=#HEX` | White bonds for dark mode or custom hex/CSS ink color |
| **Highlight** | `hl:<smarts>` | `highlight:<smarts>`, `hl=<smarts>` | Colors atoms/bonds matching a SMILES/SMARTS pattern |
| **Labels** | `num` · `stereo` · `h` · `all-c` | `idx`, `chiral`, `unfoldh`, `carbon` | Atom numbers, CIP *R/S*/*E/Z*, explicit `-H`, explicit `C` labels |
| **Transforms** | `flip` · `flipy` · `rot90`/`180`/`270` | `flipx` *(also accepted in 4th arg)* | Horizontal/vertical mirror flip and clockwise rotation |

## 7. Status & Error Code Quick Reference

| Code | Meaning | Resolution |
| :---: | :--- | :--- |
| **`#N/A`** *(~0.1s)* | Async rendering or uncached PubChem resolution in progress. | Updates automatically in a moment—no action needed. |
| **`#VALUE!`** | Empty input cell or malformed SMILES/InChI syntax. | Click **`[SMILES]` ➔ `[Diagnostics]`** for exact parser details. |
| **`#NAME?`** | Add-in inactive, or `=SMILES2IMG` typed on Excel 2016–2021. | Check [Installation](/guide/install) or append **`.FLOAT`**. |
