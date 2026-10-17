# Excel Formula Cheat Sheet

A quick, task-oriented lookup guide for all **25+ chemical structure rendering, molecular property, drug-likeness (Lipinski), and SAR screening formulas** in Excel.

---

## 1. 🖼️ Render 2D Structure Images (`...2IMG`)

Turn chemical text or names directly into high-resolution 2D structure images inside cells.

| Your Input Data | Formula (Microsoft 365 / Excel 2024+) | For Excel 2016 / 2019 / 2021 | Example |
| :--- | :--- | :--- | :--- |
| **SMILES String** | `=SMILES2IMG(cell, [bg], [options])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **Chemical Name or CAS #** | `=NAME2IMG(name_or_cas, [bg], [options])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **Reaction Scheme (KET)** | `=KET2IMG(cell, [bg], [options])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |
| **InChI String** | `=INCHI2IMG(cell, [bg], [options])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |

::: tip 💡 Background (`[bg]`) and Options (`[options]`) Are Optional!
* For standard color rendering, just pass the cell: `=SMILES2IMG(A2)`.
* To pass styling options while keeping the default background, leave the 2nd argument empty (`""`):  
  Example: `=SMILES2IMG(A2, "", "bw")` *(ACS Black & White line art)*
:::

---

## 2. 🏷️ Get Name, Formula & Molecular Weight

All property calculations (except online name lookup for unknown molecules) run in **0ms purely offline**.

| What You Want | Formula | Example Output (`A2` = Aspirin) | Description |
| :--- | :--- | :--- | :--- |
| **Substance Name** | `=SMILES2NAME(A2)` | `"Aspirin"` | Common chemical name via IUPAC / PubChem |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | `"C9H8O4"` | Standard Hill-notation formula |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `180.16` | Standard molecular weight in g/mol |
| **Monoisotopic Exact Mass** | `=SMILES2EXACTMASS(A2)` | `180.0423` | Exact monoisotopic mass in Da (for HRMS/MS) |

> 🔄 **Have an InChI string or KET reaction instead of SMILES?**  
> Simply replace the `SMILES2` prefix with `INCHI2` or `KET2`!  
> Examples: `=INCHI2MW(A2)`, `=INCHI2FORMULA(A2)`, `=KET2MW(A2)`

---

## 3. 💊 Drug-Likeness (Lipinski Ro5) & ADME Properties

Screen oral bioavailability and physicochemical descriptors directly in your spreadsheet.

| What You Want | Formula | Example Output | Description |
| :--- | :--- | :--- | :--- |
| **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `"PASS"` / `"FAIL"` | Evaluates oral drug-likeness criteria automatically |
| **LogP (Lipophilicity)** | `=SMILES2LOGP(A2)` | `1.31` | Wildman-Crippen octanol-water partition coefficient |
| **Polar Surface Area (TPSA)** | `=SMILES2TPSA(A2)` | `63.60` | Topological Polar Surface Area in Å² |
| **H-Bond Donors (HBD)** | `=SMILES2HBD(A2)` | `1` | Count of hydrogen bond donors ($-OH$, $-NH$) |
| **H-Bond Acceptors (HBA)** | `=SMILES2HBA(A2)` | `4` | Count of hydrogen bond acceptors ($N$, $O$) |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `3` | Count of rotatable single bonds |

---

## 4. 🔍 Substructure Search & Similarity (SAR Screening)

Filter compound libraries by scaffold or rank analogs against a lead molecule.

| Goal | Formula | Example Output | How It Works |
| :--- | :--- | :--- | :--- |
| **Substructure Match** | `=SMILES.MATCH(target, "query")` | `TRUE` / `FALSE` | Returns `TRUE` if `target` contains the SMILES/SMARTS substructure `query` |
| **2D Fingerprint Similarity** | `=SMILES.SIMILARITY(mol1, mol2, [metric])` | `0.0` to `1.0` | Computes 2D similarity score (`"tanimoto"` by default, or `"tversky"`) |

---

## 5. 🔄 Convert Between Name, CAS, SMILES & InChI

| Conversion | Formula | Example Output |
| :--- | :--- | :--- |
| **Name or CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` or `=NAME2SMILES("Aspirin")` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES ➔ InChI** | `=SMILES2INCHI(A2)` | `"InChI=1S/C9H8O4/c1-6(10)..."` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES ➔ InChIKey (27-char hash)** | `=SMILES2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

---

## 6. 🎨 Image Styling Option Keywords (`options`)

Combine any of these keywords in the 3rd argument of `=SMILES2IMG(A2, [bg], [options])` using spaces or commas (e.g., `"bw num stereo"`).

### Color & Highlighting Options
| Keyword | Effect | Example Formula |
| :--- | :--- | :--- |
| `colored` | Standard IUPAC element colors (default) | `=SMILES2IMG(A2, "", "colored")` |
| `bw` | High-contrast monochrome line art (ACS / patent style) | `=SMILES2IMG(A2, "", "bw")` |
| `hl:<smarts>` | Highlights matching SMILES/SMARTS substructure in color | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| `white` | White bonds and labels for dark backgrounds | `=SMILES2IMG(A2, "#1e1e1e", "white")` |
| `ink=#HEX` | Custom hex color for bonds and atom labels | `=SMILES2IMG(A2, "", "ink=#0055FF")` |

### Atom Labels & Geometry Options
| Keyword | Effect | Example Formula |
| :--- | :--- | :--- |
| `num` | Displays atom index numbers (1, 2, 3...) for NMR/MS | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | Displays CIP stereochemistry labels (*R/S*, *E/Z*) | `=SMILES2IMG(A2, "", "stereo")` |
| `h` | Unfolds and shows all explicit hydrogen atoms ($H$) | `=SMILES2IMG(A2, "", "h")` |
| `all-c` | Shows explicit `C` symbols on all carbon vertices | `=SMILES2IMG(A2, "", "all-c")` |
| `flip` / `flipv` | Flips molecule horizontally (`flip`) or vertically (`flipv`) | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` / `rot270` | Rotates molecule clockwise by 90°, 180°, or 270° | `=SMILES2IMG(A2, "", "rot90")` |

---

## 7. 📋 Copy-and-Paste Formula Recipes

```text
# 1. Render directly from a CAS Registry Number or common name
=NAME2IMG("50-78-2")
=NAME2IMG("Caffeine")

# 2. Highlight a benzene ring (c1ccccc1) inside a molecule
=SMILES2IMG(A2, "", "hl:c1ccccc1")

# 3. ACS monochrome line art with R/S stereochemistry labels
=SMILES2IMG(A2, "", "bw stereo")

# 4. Dark Mode cell (dark background + crisp white bonds)
=SMILES2IMG(A2, "#1e1e1e", "white")

# 5. Check if a molecule contains a carboxylic acid group (TRUE / FALSE)
=SMILES.MATCH(A2, "C(=O)O")

# 6. Calculate 2D Tanimoto similarity against a lead compound in $B$1
=SMILES.SIMILARITY(A2, $B$1)
```
