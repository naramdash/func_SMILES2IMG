# Excel Chemical Formula Cheat Sheet

Quick-lookup reference tables for all **25+ Excel chemical functions and rendering options**, organized by task.

> 🛠️ **Want to build and preview formulas interactively by clicking options?**  
> 👉 Open the [**🛠️ Excel Formula Generator & Style Simulator ➔**](/tools/formula-builder) at the bottom of the sidebar!

---

## 1. 🖼️ Render 2D Structure Images (`...2IMG`)

Embed high-resolution 2D chemical structures or multi-step reaction schemes directly into cells.

| Your Input Data | Formula (Microsoft 365 / Excel 2024+) | For Excel 2016 / 2019 / 2021 | Example |
| :--- | :--- | :--- | :--- |
| **SMILES String** | `=SMILES2IMG(cell, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **Chemical Name or CAS #** | `=NAME2IMG(name_or_cas, [bg], [style], [transform])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **Reaction Scheme (KET)** | `=KET2IMG(cell, [bg], [style], [transform])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |
| **InChI String** | `=INCHI2IMG(cell, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |

::: tip 💡 Shorthand Syntax Tips
* **Default Color + Transparent BG:** Only the 1st argument is required: `=SMILES2IMG(A2)`.
* **Quick B&W Shortcut:** Pass `"bw"` (or `FALSE`) directly in the 2nd argument to get monochrome line art with a transparent background (`=SMILES2IMG(A2, "bw")`).
* **Flexible Delimiters:** Combine multiple style tokens in the 3rd argument using spaces (` `), commas (`,`), pipes (`|`), semicolons (`;`), or plus signs (`+`)—e.g., `"bw num stereo"` or `"bw, num, stereo"`.
:::

---

## 2. 🏷️ Get Name, Formula, MW & Exact Mass

Except for online common-name lookup (`=SMILES2NAME`), all property functions run in **0ms purely offline**.

| What You Want | SMILES Formula | For InChI / KET Input | Example Output (`A2` = Aspirin) | Description |
| :--- | :--- | :--- | :--- | :--- |
| **Substance Name** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)`<br>`=KET2NAME(A2)` | `"Aspirin"` | Common name via IUPAC / PubChem (with local disk cache) |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)`<br>`=KET2FORMULA(A2)` | `"C9H8O4"` | Standard Hill-notation molecular formula |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)`<br>`=KET2MW(A2)` | `180.16` | Standard average molecular weight in g/mol (2 decimals) |
| **Monoisotopic Exact Mass** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)`<br>`=KET2EXACTMASS(A2)` | `180.0423` | Exact monoisotopic mass in Da for HRMS/LC-MS (4 decimals) |

---

## 3. 💊 Drug-Likeness (Lipinski Ro5) & ADME Properties

Evaluate oral bioavailability and physicochemical descriptors directly in your spreadsheet. Accepts SMILES, InChI, or KET inputs automatically.

| What You Want | Formula | Example Output | Lipinski Cutoff & Description |
| :--- | :--- | :--- | :--- |
| **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `"PASS"` / `"FAIL"` | Evaluates all 4 rules (MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10) |
| **LogP (Lipophilicity)** | `=SMILES2LOGP(A2)` | `1.31` | Wildman-Crippen octanol-water partition coefficient (≤ 5.0) |
| **Polar Surface Area (TPSA)** | `=SMILES2TPSA(A2)` | `63.60` | Topological Polar Surface Area in Å² (≤ 140 oral, ≤ 90 CNS/BBB) |
| **H-Bond Donors (HBD)** | `=SMILES2HBD(A2)` | `1` | Count of `-OH` and `-NH` hydrogen bond donors (≤ 5) |
| **H-Bond Acceptors (HBA)** | `=SMILES2HBA(A2)` | `4` | Count of nitrogen (`N`) and oxygen (`O`) acceptors (≤ 10) |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `3` | Count of rotatable single bonds (≤ 10 recommended) |

---

## 4. 🔍 Substructure Search & Similarity (SAR Screening)

Filter compound libraries by scaffold or rank analogs against a lead molecule. For a copy-and-paste dictionary of 12 common SMARTS functional groups, see the **[Excel SAR & SDF Guide](/guide/excel-cheminformatics)**.

| Goal | Formula | Example Output | How It Works |
| :--- | :--- | :--- | :--- |
| **Substructure Match** | `=SMILES.MATCH(target, "query")` | `TRUE` / `FALSE` | Returns `TRUE` if `target` contains the SMILES/SMARTS substructure `query` |
| **2D Fingerprint Similarity** | `=SMILES.SIMILARITY(mol1, mol2, [metric])` | `0.0` to `1.0` | Computes 2D similarity score (`"tanimoto"` default, `"tversky"`, or `"euclid-sub"`) |

---

## 5. 🔄 Convert Between Name, CAS, SMILES & InChI

| Conversion | Formula | Example Output |
| :--- | :--- | :--- |
| **Name or CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` or `=NAME2SMILES("Aspirin")` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` / `=KET2INCHI(A2)` | `"InChI=1S/C9H8O4/c1-6(10)..."` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` / `=INCHI2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

---

## 6. 🎨 Complete Background (`[bg]`) & Style (`[style]`) Reference

### 2nd Argument: Background (`[bg]`)
| Value | Effect | Example |
| :--- | :--- | :--- |
| `""` or `"trans"` / `"transparent"` | Transparent alpha background (default) | `=SMILES2IMG(A2, "")` |
| `"white"` | Solid opaque white background box | `=SMILES2IMG(A2, "white")` |
| `"#RRGGBB"` or CSS color name | Custom background fill (`"#1e1e1e"`, `"yellow"`, etc.) | `=SMILES2IMG(A2, "#1e1e1e", "white")` |

### 3rd Argument: Style, Label & Geometry Tokens (`[style]`)
| Category | Primary Token | Accepted Aliases | Effect |
| :--- | :--- | :--- | :--- |
| **Color Mode** | `color` | `colour`, `cpk`, `true` | Standard IUPAC element colors (default) |
| **Monochrome** | `bw` | `mono`, `monochrome`, `black`, `gray`, `grayscale` | High-contrast ACS journal & patent black-and-white line art |
| **Bond Ink Color** | `white` / `ink=#HEX` | `light`, `line=#HEX`, `base=#HEX` | Crisp white bonds for dark mode, or **any custom brand/slide ink color** (`"ink=#1e40af"`, `"line=navy"`) |
| **Substructure Highlight** | `hl:<smarts>` | `highlight:<smarts>`, `hl=<smarts>`, `highlight=<smarts>` | Highlights matching SMILES/SMARTS substructure in color |
| **Atom Numbers** | `num` | `number`, `numbers`, `idx`, `index`, `atom-num`, `atom-ids` | Displays 1-based atom index numbers for NMR/MS |
| **Stereochemistry** | `stereo` | `chiral`, `abs`, `ext`, `stereochemistry` | Displays CIP stereochemistry labels (*R/S*, *E/Z*) |
| **Explicit Hydrogens** | `h` | `hydrogen`, `hydrogens`, `unfoldh`, `all-h` | Unfolds and displays all explicit hydrogen atoms (`H`) |
| **Carbon Symbols** | `all-c` | `carbon`, `carbons`, `c` | Explicitly displays `C` symbols on all carbon vertices |
| **Mirror Flip** | `flip` / `flipy` | `flipx` (horizontal), `flipy` (vertical) | Mirrors molecular coordinates horizontally (`flip`) or vertically (`flipy`) |
| **Rotation** | `rot90` / `rot180` / `rot270` | *(Also accepted in 4th `[transform]` argument)* | Rotates structure clockwise by 90°, 180°, or 270° |

---

## 7. 📊 Return Data Types & Excel Error Code Reference

### Return Data Types (Formula Math & Logic Compatibility)
| Function Group | Excel Return Type | Formula & Math Usage Example |
| :--- | :--- | :--- |
| `=SMILES2MW`, `=SMILES2EXACTMASS`, `=SMILES2LOGP`, `=SMILES2TPSA`, `=SMILES.SIMILARITY` | **Numeric (`Double`)** | `=SMILES2MW(A2) * 5.0` (stoichiometry mass), `=SMILES2TPSA(A2) <= 90` |
| `=SMILES2HBD`, `=SMILES2HBA`, `=SMILES2ROTBONDS` | **Numeric (`Integer`)** | `=SMILES2HBD(A2) + SMILES2HBA(A2)` |
| `=SMILES.MATCH` | **Boolean (`TRUE / FALSE`)** | `=IF(SMILES.MATCH(A2, "c1ccccc1"), "Aromatic", "Aliphatic")` |
| `=SMILES2LIPINSKI`, `=SMILES2FORMULA`, `=SMILES2NAME`, `=NAME2SMILES`, `=SMILES2INCHI` | **Text (`String`)** | `=IF(SMILES2LIPINSKI(A2)="PASS", "Candidate", "Exclude")` |
| `=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG` | **In-Cell Picture (`.FLOAT` = Shape)** | Scales automatically with cell height/width and sorts with table rows |

### Understanding Excel Status & Error Codes (`#N/A`, `#VALUE!`, `#NAME?`)
| Displayed Code | Meaning & Cause | How to Resolve |
| :---: | :--- | :--- |
| **`#N/A`** *(briefly ~0.1s)* | **Normal Async Rendering in Progress:** Shown momentarily (~0.1s) while generating a new structure image or resolving an uncached name from PubChem, then **automatically updates to the final image or value**. | Wait a fraction of a second—no action needed. |
| **`#VALUE!`** | **Invalid Chemical Syntax or Empty Reference:** The referenced cell is empty, or the SMILES/InChI string contains a chemical syntax error (e.g., unclosed ring number or parenthesis). | Click **`[SMILES]`** ➔ **`[Diagnostics]`** on the ribbon to view the exact parser error message. |
| **`#NAME?`** | **Add-In Not Loaded or Excel Version Mismatch:** Either the add-in is not active, or `=SMILES2IMG` was typed on Excel 2016–2021 instead of `=SMILES2IMG.FLOAT`. | Run the [10-Second Quick Install](/guide/getting-started), or append `.FLOAT` on Excel 2016–2021. |

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../tools/formula-builder" class="s2i-card">
    <div class="s2i-card-title">🛠️ Formula Generator & Style Simulator</div>
    <p class="s2i-card-desc">Click options to preview rendered structures in real time and copy the formula</p>
  </a>
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel Cheminformatics: SAR & SDF</div>
    <p class="s2i-card-desc">12 common SMARTS functional groups, HRMS adduct formulas, and Excel IF/Filter recipes</p>
  </a>
</div>


