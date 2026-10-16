# Excel Function & Parameter Specification

Complete reference for Excel worksheet functions provided by **SMILES for Microsoft Office Suite**.

---

## 📐 Function Signatures

### 1. `SMILES2IMG`
```text
=SMILES2IMG(smiles_or_ket, [background], [options])
```
* **Purpose:** Renders chemical structure as a **native in-cell image** (Microsoft 365 / Excel 2024+).
* **Return Value:** In-cell image entity that auto-scales with cell boundaries.

### 2. `SMILES2IMG.FLOAT`
```text
=SMILES2IMG.FLOAT(smiles_or_ket, [background], [options])
```
* **Purpose:** Inserts a floating shape locked to cell bounds (`xlMoveAndSize`). Compatible with Excel 2016, 2019, and 2021.
* **Return Value:** In-cell text string while managing the associated overlay shape.

### 3. `KET2IMG` & `KET2IMG.FLOAT`
```text
=KET2IMG(ket_json_string, [background], [options])
=KET2IMG.FLOAT(ket_json_string, [background], [options])
```
* **Purpose:** Renders multi-component organic chemical reactions, reaction arrows, plus signs, and text annotations directly from Ketcher KET JSON representations.

### 4. `INCHI2IMG` & `INCHI2IMG.FLOAT`
```text
=INCHI2IMG(inchi_string, [background], [options], [transform])
=INCHI2IMG.FLOAT(inchi_string, [background], [options], [transform])
```
* **Purpose:** Renders chemical structures directly from IUPAC standard InChI strings (both `InChI=1S/...` and prefix-less `1S/...`).

### 5. `NAME2IMG`, `NAME2IMG.FLOAT`, & `NAME2SMILES` *(New in v1.7.0)*
```text
=NAME2IMG(name_or_cas, [background], [options], [transform])
=NAME2IMG.FLOAT(name_or_cas, [background], [options], [transform])
=NAME2SMILES(name_or_cas)
```
* **Purpose:** Resolves common chemical names (e.g. `"Aspirin"`, `"Caffeine"`), IUPAC names, or CAS Registry Numbers (e.g. `"50-78-2"`) to canonical SMILES or directly to an image (offline dictionary for 60+ common compounds + PubChem REST API fallback with persistent disk caching).

---

## 🔬 Molecular Property & Drug-Likeness Functions

These functions compute chemical and physicochemical properties directly from SMILES, Ketcher KET, or IUPAC InChI strings:

| Function | Return Value | Description | Example Output |
| :--- | :--- | :--- | :--- |
| `=SMILES2NAME(smiles)` / `=KET2NAME(ket)` / `=INCHI2NAME(inchi)` | `string` | Common Substance Name (via PubChem / IUPAC) | `"Aspirin"`, `"Caffeine"` |
| `=NAME2SMILES(name_or_cas)` | `string` | Resolves chemical name or CAS Registry Number to SMILES | `"CC(=O)Oc1ccccc1C(=O)O"` |
| `=SMILES2MW(smiles)` / `=KET2MW(ket)` / `=INCHI2MW(inchi)` | `double` | Molecular Weight in g/mol (Pure local calculation) | `180.16`, `194.19` |
| `=SMILES2EXACTMASS(smiles)` / `=KET2EXACTMASS(ket)` / `=INCHI2EXACTMASS(inchi)` | `double` | Monoisotopic Exact Mass in Da for HRMS/MS (Pure local) | `180.0423` |
| `=SMILES2FORMULA(smiles)` / `=KET2FORMULA(ket)` / `=INCHI2FORMULA(inchi)` | `string` | Standard Hill-notation chemical formula (Pure local calculation) | `"C9H8O4"`, `"C8H10N4O2"` |
| `=SMILES2INCHI(smiles)` / `=KET2INCHI(ket)` | `string` | Standard InChI identifier (Pure local calculation) | `"InChI=1S/C9H8O4/c1-6(10)13-8-5..."` |
| `=INCHI2SMILES(inchi)` | `string` | Canonical SMILES converted from InChI (Pure local calculation) | `"CC(=O)Oc1ccccc1C(=O)O"` |
| `=SMILES2INCHIKEY(smiles)` / `=KET2INCHIKEY(ket)` / `=INCHI2INCHIKEY(inchi)` | `string` | 27-character standard InChIKey hash (Pure local calculation) | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |
| `=SMILES2LOGP(structure)` | `double` | Calculated Wildman-Crippen octanol-water LogP (Pure local) | `1.31` |
| `=SMILES2TPSA(structure)` | `double` | Topological Polar Surface Area in Å² (Pure local) | `63.60` |
| `=SMILES2HBD(structure)` | `int` | Hydrogen Bond Donor count (OH + NH, Pure local) | `1` |
| `=SMILES2HBA(structure)` | `int` | Hydrogen Bond Acceptor count (N + O, Pure local) | `4` |
| `=SMILES2ROTBONDS(structure)` | `int` | Rotatable Bond count (Pure local) | `3` |
| `=SMILES2LIPINSKI(structure)` | `string` | Evaluates Lipinski's Rule of 5 and returns `"PASS"` or `"FAIL"` | `"PASS"` |

---

## 🧪 Cheminformatics & SAR Screening Functions *(New in v1.7.0)*

| Function | Return Value | Description | Example |
| :--- | :--- | :--- | :--- |
| `=SMILES.MATCH(target, query)` | `bool` | Returns `TRUE` if `target` contains the SMILES/SMARTS substructure `query` | `=SMILES.MATCH(A2, "c1ccccc1")` |
| `=SMILES.SIMILARITY(mol1, mol2, [metric])` | `double` | Calculates 2D fingerprint similarity (`0.0` to `1.0`). Metric: `"tanimoto"` (default) or `"tversky"` | `=SMILES.SIMILARITY(A2, $B$1, "tanimoto")` |

> 💡 **Tip:** Structural property and SAR functions run in 0ms purely offline. Combine `=SMILES.MATCH`, `=SMILES.SIMILARITY`, and `=SMILES2LIPINSKI` with Excel filters to screen thousands of compounds directly in your spreadsheet!

---

## ⚙️ Parameter Reference

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `smiles_or_ket` | `string` | *(Required)* | Valid SMILES string (e.g., `CN1C=NC2=C1C(=O)N(C(=O)N2C)C`) or Ketcher JSON string (`{"root": ...}`). |
| `background` | `string` | `""` (Auto) | Background color. Empty string defaults to transparent in Ketcher and transparent/white in Excel. Accepts hex codes (`"#FFFFFF"`, `"#1A1A1A"`, `"#00000000"`), named colors (`"white"`, `"transparent"`, `"black"`). |
| `options` | `string` | `""` | Space-, comma-, semicolon-, or pipe-separated styling and rendering flags. |

---

## 🎨 Option Tokens Specification

| Token | Category | Description | Example |
| :--- | :--- | :--- | :--- |
| `bw` / `monochrome` | Mode | Black-and-white line art (ACS / patent specification) | `"bw"` |
| `colored` | Mode | Standard IUPAC CPK element color scheme | `"colored"` |
| `hl:<smarts>` / `highlight:<smarts>` | Highlight | Highlights atoms and bonds matching a SMILES/SMARTS substructure query | `"hl:c1ccccc1"` |
| `num` / `numbers` | Label | Display 1-based atom indices for NMR/MS assignments | `"num"` |
| `stereo` | Label | Force CIP stereochemistry annotations (*R/S*, *E/Z*) | `"stereo"` |
| `h` / `hydrogens` | Label | Unfold and display explicit hydrogen atoms | `"h"` |
| `all-c` / `carbons` | Label | Explicitly display all carbon atom symbols (`C`) | `"all-c"` |
| `white` / `ink=#HEX` | Mode | Custom bond/atom ink color (e.g. `"white"` for dark backgrounds) | `"white"` |
| `flip` | Transform | Flips molecular coordinates horizontally | `"flip"` |
| `flipv` | Transform | Flips molecular coordinates vertically | `"flipv"` |
| `rot90` | Transform | Rotates molecule 90 degrees clockwise | `"rot90"` |
| `rot180` | Transform | Rotates molecule 180 degrees | `"rot180"` |
| `rot270` | Transform | Rotates molecule 270 degrees clockwise | `"rot270"` |

---

## 🔀 Flexible Delimiter Syntax

Option tokens can be formatted with any standard delimiter:

```text
=SMILES2IMG(A2, "", "bw num stereo")
=SMILES2IMG(A2, "", "bw, num, stereo")
=SMILES2IMG(A2, "", "bw; num; stereo")
=SMILES2IMG(A2, "", "bw | num | stereo")
```

---

## 💡 Practical Examples

* **Substructure Highlighting (Scaffold / Pharmacophore):**
  ```text
  =SMILES2IMG(A2, "trans", "hl:c1ccccc1")
  ```
* **Direct CAS / Chemical Name Rendering:**
  ```text
  =NAME2IMG("50-78-2")
  =NAME2IMG("Ibuprofen", "white")
  ```
* **In-Cell SAR Substructure Screening:**
  ```text
  =SMILES.MATCH(A2, "c1ccccc1")
  =SMILES.MATCH(A2, "[#6]~[#7]")
  ```
* **Hit Compound 2D Fingerprint Similarity:**
  ```text
  =SMILES.SIMILARITY(A2, $B$1, "tanimoto")
  ```
* **Drug-Likeness (Lipinski Rule of 5) Oral Bioavailability Evaluation:**
  ```text
  =SMILES2LIPINSKI(A2)
  ```
* **Monochrome Patent Line Art:**
  ```text
  =SMILES2IMG(A2, "", "bw")
  ```
* **Dark Mode with White Bonds:**
  ```text
  =SMILES2IMG(A2, "#1e1e1e", "white")
  ```
* **NMR Peak Mapping with Stereochemistry:**
  ```text
  =SMILES2IMG(A2, "", "num stereo")
  ```
* **Full Multi-Step Organic Reaction Scheme:**
  ```text
  =KET2IMG(B2)
  ```
