# All 49 Excel Functions

Every **`func_SMILES2IMG`** formula follows a unified **`=<INPUT>2<OUTPUT>`** design: pick a function prefix matching your **input** (`SMILES`, `NAME`, `INCHI`, or `KET`) and a suffix matching the **cell output** you need—with **full parity across `SMILES`, `INCHI`, and `KET` inputs**.

::: tip Two Shared Pipelines Across All 49 Functions
* **Structure Image Output (`8` functions):** `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, and `=KET2IMG` (plus their `.FLOAT` equivalents) share the identical 4-argument signature `(input, [bg], [style], [transform])` and the same pure EPAM Indigo renderer.
* **Scalar Output (`41` functions):** Property, 3D complexity, SMARTS substructure match, and identifier conversion functions return native Excel numbers, booleans (`TRUE`/`FALSE`), or strings—ready for `IF`, `FILTER`, `SORT`, `XLOOKUP`, and stoichiometry math (`MW × mmol`). `INCHI2*` and `KET2*` strictly validate their respective format inputs (`#VALUE!` on mismatch).
:::

## ⭐ Top 5 Ready-to-Copy Formulas

Click **Copy All** and paste into cell **`A1`** to populate inputs (`A2:A5`) alongside live formulas (`B2:B6`), or copy any row individually (on Excel 2016–2021, append **`.FLOAT`** to `...2IMG` functions):

<FormulaTable
  title="Top 5 Ready-to-Copy Lab Sheet"
  :columns='["SMILES / Input", "Formula"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG(A2)", desc: "2D IUPAC color structure from SMILES in A2", out: "2D Ethanol" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A3, \"\", \"bw\")", desc: "ACS journal monochrome B&W structure from A3", out: "B&W Aspirin" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A4)", desc: "Molecular weight in g/mol (offline)", out: "180.16" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2TPSA(A5)", desc: "Topological Polar Surface Area (offline)", out: "63.60" },
    { input: "", formula: "=NAME2IMG(\"Aspirin\")", desc: "Standalone 2D structure from name or CAS #", out: "2D Aspirin" }
  ]'
/>

## 🗺️ Guide Map by Category

| Sidebar Category | Included Functions (`49` Total) | Pages |
| :--- | :--- | :--- |
| **Structure Images (`8`)** | `SMILES2IMG`, `NAME2IMG`, `INCHI2IMG`, `KET2IMG` (+ 4 `.FLOAT`) | [SMILES, InChI](/excel/structures) · [Names, Solvents, CAS](/excel/name-cas) · [Reaction Schemes (KET)](/excel/reactions) · [Styles & Highlights](/excel/styles-reactions) |
| **Core & Name (`12`)** | `SMILES2*`, `INCHI2*`, `KET2*` × (`NAME`, `FORMULA`, `MW`, `EXACTMASS`) | [Formula, MW, Exact Mass](/excel/properties) |
| **LogP, TPSA & 3D (`18`)** | `SMILES2*`, `INCHI2*`, `KET2*` × (`LOGP`, `TPSA`, `HEAVY`, `RINGS`, `AROMATIC`, `FSP3`) | [LogP, TPSA & 3D](/excel/lipinski-adme) |
| **SAR Search (`3`)** | `SMILES.MATCH`, `INCHI.MATCH`, `KET.MATCH` | [Substructure Search](/excel/sar) |
| **Identifier Conversions (`8`)** | `NAME2SMILES`, `KET2SMILES`, `INCHI2SMILES`, `*2INCHI` (×2), `*2INCHIKEY` (×3) | [Identifier Conversion](/excel/data-sdf) |

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

Except for online common-name lookups (`*2NAME`, cached locally after 1 query), all property calculations run **locally offline** ([Formula, MW, Exact Mass](/excel/properties)):

| Output | SMILES Input | InChI Input | KET Input | Type | Aspirin (`A2`) |
| :--- | :--- | :--- | :--- | :---: | :--- |
| **Substance Name** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)` | `=KET2NAME(A2)` | `String` | `"Aspirin"` |
| **Chemical Formula (Hill)** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)` | `=KET2FORMULA(A2)` | `String` | `"C9H8O4"` |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)` | `=KET2MW(A2)` | `Double` | `180.16` |
| **Monoisotopic Exact Mass** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)` | `=KET2EXACTMASS(A2)` | `Double` | `180.0423` |

## 3. LogP, TPSA & 3D Structural Complexity (`18`)

Full parity across SMILES, InChI, and KET inputs ([LogP, TPSA & 3D](/excel/lipinski-adme)):

| Descriptor | SMILES Formula | InChI Formula | KET Formula | Type | Aspirin |
| :--- | :--- | :--- | :--- | :---: | :---: |
| **LogP (Crippen)** | `=SMILES2LOGP(A2)` | `=INCHI2LOGP(A2)` | `=KET2LOGP(A2)` | `Double` | `1.31` |
| **TPSA (Å²)** | `=SMILES2TPSA(A2, [include_sp])` | `=INCHI2TPSA(A2, [include_sp])` | `=KET2TPSA(A2, [include_sp])` | `Double` | `63.60` |
| **Heavy Atoms** | `=SMILES2HEAVY(A2)` | `=INCHI2HEAVY(A2)` | `=KET2HEAVY(A2)` | `Integer` | `13` |
| **Total Rings (SSSR)** | `=SMILES2RINGS(A2)` | `=INCHI2RINGS(A2)` | `=KET2RINGS(A2)` | `Integer` | `1` |
| **Aromatic Rings** | `=SMILES2AROMATIC(A2)` | `=INCHI2AROMATIC(A2)` | `=KET2AROMATIC(A2)` | `Integer` | `1` |
| **Fraction Csp³ (Fsp³)** | `=SMILES2FSP3(A2)` | `=INCHI2FSP3(A2)` | `=KET2FSP3(A2)` | `Double` | `0.11` |

## 4. Substructure Search (`3`)

All three notations support SMILES and SMARTS substructure queries directly ([Substructure Search](/excel/sar)):

| Output | SMILES Formula | InChI / KET Parity | Type | Result |
| :--- | :--- | :--- | :---: | :---: |
| **Substructure Match** | `=SMILES.MATCH(target, "query")` | `=INCHI.MATCH(inchi, "query")` · `=KET.MATCH(ket, "query")` | `Boolean` | `TRUE` / `FALSE` |

## 5. Identifier Conversions (`8`)

Standardize Chemical Name, CAS #, Canonical SMILES, KET JSON, IUPAC InChI, and 27-character InChIKey ([Identifier Conversion](/excel/data-sdf)):

| Conversion | Formula | Type | Example Output |
| :--- | :--- | :---: | :--- |
| **Name or CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **InChI ➔ Canonical SMILES** | `=INCHI2SMILES(A2)` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **KET ➔ Canonical SMILES** | `=KET2SMILES(A2)` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` · `=KET2INCHI(A2)` | `String` | `"InChI=1S/C9H8O4/..."` |
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
| **Rotations** | `rot90` · `rot180` · `rot270` | *(accepted in 3rd or 4th arg)* | Rigid planar clockwise rotations (preserves stereochemistry) |

## 7. Status & Error Code Quick Reference

| Code | Meaning | Resolution |
| :---: | :--- | :--- |
| **`#N/A`** | Offline network, compound name not found on PubChem (`*2NAME`, `NAME2*`), or uncomputable property/conversion. | Check network connectivity, compound spelling, or input type (e.g., reaction or R-group `*` passed to single-molecule property). *(`...2IMG` formulas stay briefly blank while queued rendering completes.)* |
| **`#VALUE!`** | Syntax error, non-InChI passed to `INCHI2*`, non-KET passed to `KET2*`, or invalid formula syntax. | Open **`[SMILES]` ➔ `[Diagnostics]`** to inspect the exact EPAM Indigo parser message. |
