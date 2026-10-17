# Excel Cheminformatics: Properties, SAR Screening & SDF Files

Go beyond structure rendering: learn how to automatically calculate **Molecular Weight, Exact Mass, and Formulas**, screen compounds for **Lipinski Rule of 5 drug-likeness**, filter libraries by **SMARTS substructures & 2D similarity**, and import/export **`.sdf` database files** in Excel.

::: tip ⚡ 100% Offline · Zero Latency (0ms)
Except for online common-name lookup (`=SMILES2NAME`), all physicochemical calculations, substructure searches, and similarity scores run **locally on your CPU in 0ms** without network calls. You can sort and filter thousands of compounds in real time.
:::

---

## 1. Auto-Calculate Basic Properties (Name · Formula · MW · Exact Mass)

When cell `A2` contains a SMILES string (e.g., Aspirin `CC(=O)Oc1ccccc1C(=O)O`), place these formulas in adjacent columns:

| Property | Excel Formula | Example Output (Aspirin) | Practical Tip |
| :--- | :--- | :--- | :--- |
| **Substance Name** | `=SMILES2NAME(A2)` | `"Aspirin"` | Uses built-in 60+ reagent dictionary + PubChem lookup with disk caching |
| **Name / CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` | `"CC(=O)Oc1ccccc1C(=O)O"` | Converts common names or CAS Registry Numbers into canonical SMILES |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | `"C9H8O4"` | Standard Hill-notation molecular formula |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `180.16` | Standard average molecular weight (g/mol) for stoichiometry & molarity |
| **Monoisotopic Exact Mass** | `=SMILES2EXACTMASS(A2)` | `180.0423` | **For HRMS / LC-MS peak matching** (monoisotopic mass in Da, 4 decimals) |
| **Standard InChIKey** | `=SMILES2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-..."` | 27-character hash—ideal for deduplicating compounds across sheets (`VLOOKUP`/`XLOOKUP`) |

::: info 🧪 LC-MS / HRMS Tip: Calculating Adduct m/z Values
High-resolution mass spectrometry measures monoisotopic mass (`=SMILES2EXACTMASS(A2)`) rather than average molecular weight (`MW`). Combine it with standard adduct masses directly in Excel:
* **`[M+H]⁺` (Protonated):** `=SMILES2EXACTMASS(A2) + 1.0073`
* **`[M+Na]⁺` (Sodiated):** `=SMILES2EXACTMASS(A2) + 22.9892`
* **`[M−H]⁻` (Deprotonated):** `=SMILES2EXACTMASS(A2) - 1.0073`
:::

> 🔄 **Working with InChI or KET Reactions?** Simply replace the `SMILES2` prefix with `INCHI2` or `KET2` (e.g., `=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).

---

## 2. Oral Drug-Likeness (Lipinski Rule of 5) & ADME Descriptors

Evaluate whether candidate compounds have favorable physicochemical properties for oral bioavailability directly inside your worksheet.

### Automatic Lipinski Rule of 5 Check (`=SMILES2LIPINSKI`)
```text
=SMILES2LIPINSKI(A2)
```
Evaluates the 4 classic Lipinski criteria and returns **`"PASS"`** or **`"FAIL"`**:

| Descriptor | Individual Formula | Lipinski Cutoff | Aspirin Example | Medicinal Chemistry Role |
| :--- | :--- | :---: | :---: | :--- |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | Size limit for passive membrane permeability |
| **Lipophilicity (LogP)** | `=SMILES2LOGP(A2)` | **≤ 5.0** | `1.31` | Wildman-Crippen octanol/water partition balance |
| **H-Bond Donors (HBD)** | `=SMILES2HBD(A2)` | **≤ 5** | `1` | Total count of `-OH` and `-NH` hydrogens |
| **H-Bond Acceptors (HBA)** | `=SMILES2HBA(A2)` | **≤ 10** | `4` | Total count of nitrogen (`N`) and oxygen (`O`) atoms |

### Additional ADME Descriptors: Polar Surface Area (`TPSA`) & Rotatable Bonds (`ROTBONDS`)
| Descriptor | Formula | Example | Medicinal Chemistry Rule of Thumb |
| :--- | :--- | :---: | :--- |
| **Topological Polar Surface Area** | `=SMILES2TPSA(A2)` | `63.60` (Å²) | • **≤ 140 Å²:** Good oral intestinal absorption<br>• **≤ 90 Å²:** Favorable for CNS Blood-Brain Barrier (BBB) penetration |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `3` | • **≤ 10 bonds (Veber's Rule):** Favorable conformational flexibility for oral bioavailability |

---

## 3. Substructure Search (`=SMILES.MATCH`) & SMARTS Dictionary

Screen thousands of rows to find molecules containing a specific scaffold or functional group:

```text
=SMILES.MATCH(A2, "query_pattern")
```
* Returns **`TRUE`** if the substructure is present and **`FALSE`** otherwise—making it effortless to filter your table with Excel's **AutoFilter**.
* Pass the exact same pattern into `=SMILES2IMG(A2, "", "hl:pattern")` to visually highlight the matching atoms and bonds!

### 📋 Copy-and-Paste Dictionary of 12 Common Functional Groups (SMILES / SMARTS)
No need to memorize SMARTS syntax—copy any pattern below straight into your formula:

| Scaffold / Functional Group | Pattern Code | Match Formula Example | Highlight Image Formula |
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

---

## 4. Rank Analogs by 2D Fingerprint Similarity (`=SMILES.SIMILARITY`)

Compare every compound in your sheet (`A2`, `A3`...) against a reference lead compound (e.g., locked in cell `$B$1`) on a **`0.0` to `1.0` scale**:

```text
=SMILES.SIMILARITY(A2, $B$1)
```
> 💡 **Pro Tip:** Lock the reference cell with `$B$1`, drag the formula down, and sort the column **Largest to Smallest** to immediately surface the closest structural analogs!

### Choosing a Similarity Metric (`[metric]` argument)
| Metric Option | Formula Example | When to Use |
| :--- | :--- | :--- |
| **`"tanimoto"`** *(Default)* | `=SMILES.SIMILARITY(A2, $B$1, "tanimoto")` | Industry-standard balanced 2D fingerprint similarity between two molecules |
| **`"tversky"`** | `=SMILES.SIMILARITY(A2, $B$1, "tversky")` | Asymmetric similarity—ideal when comparing a smaller fragment against larger molecules |
| **`"euclid-sub"`** | `=SMILES.SIMILARITY(A2, $B$1, "euclid-sub")` | Substructure-weighted Euclidean similarity score |

---

## 5. Import & Export `.sdf` / `.mol` Chemical Database Files

Open and save multi-compound **`.sdf` (Structure-Data File)** and **`.mol`** libraries directly in Excel without external conversion scripts.

### 📥 Import an `.sdf` File into an 8-Column Medicinal Chemistry Table (`Import SDF File (.sdf)...`)
1. Click **`[SMILES] ➔ [SDF File] ➔ [Import SDF File (.sdf)...]`** on the Excel ribbon.
2. Select any `.sdf` or `.mol` file on your PC. A formatted **8-column medicinal chemistry table** (with auto-adjusted row heights and column widths) is generated automatically:

| Column | Header | Auto-Generated Content (Row 2 Example) |
| :---: | :--- | :--- |
| **Col A** | **Name** | Compound title or `<NAME>` / `<TITLE>` SD property tag |
| **Col B** | **Structure** | `=SMILES2IMG(C2)` *(or `=SMILES2IMG.FLOAT(C2)` automatically on Excel 2016–2021)* |
| **Col C** | **SMILES** | Canonical SMILES string |
| **Col D** | **Formula** | `=SMILES2FORMULA(C2)` (Standard Hill formula) |
| **Col E** | **MW (g/mol)** | `=SMILES2MW(C2)` (Average molecular weight) |
| **Col F** | **Exact Mass** | `=SMILES2EXACTMASS(C2)` (HRMS monoisotopic mass) |
| **Col G** | **LogP** | `=SMILES2LOGP(C2)` (Crippen octanol/water LogP) |
| **Col H** | **Lipinski** | `=SMILES2LIPINSKI(C2)` (`"PASS"` / `"FAIL"` Rule of 5 check) |

### 📤 Export a Worksheet Back to `.sdf` (`Export Worksheet to SDF...`)
1. Click **`[SMILES]`** ➔ **`[SDF File]`** ➔ **`[Export Worksheet to SDF...]`** and choose a destination file name.
2. Without requiring manual column mapping, the add-in automatically scans the active worksheet for compound names and SMILES/structure formulas and writes a standard V2000 `.sdf` file—ready for molecular docking or sharing with collaborators.

---

## 6. Combining with Native Excel Functions (`IF`, `AND`) & Conditional Formatting

Because all chemical formulas return native Excel numbers (`double`), booleans (`TRUE`/`FALSE`), and text, you can combine them directly with Excel's built-in logic and **Conditional Formatting**:

### ① Multi-Criteria Lead Selection (`=IF` + `AND`)
| Screening Goal | Ready-to-Copy Excel Formula (`A2` = SMILES) |
| :--- | :--- |
| **Flag hits passing Lipinski Ro5 AND containing a pyridine ring** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "✅ Priority Hit", "Exclude")` |
| **Flag CNS / BBB penetrant candidates (`TPSA ≤ 90` & `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "🧠 CNS Candidate", "Standard")` |
| **Conditionally highlight halogens only when present** | `=IF(SMILES.MATCH(A2, "[F,Cl,Br,I]"), SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]"), SMILES2IMG(A2))` |

### ② 10-Second Conditional Formatting (`PASS`/`FAIL` Badges & Similarity Data Bars)
1. **Color-Code `PASS` / `FAIL`:** Select your `=SMILES2LIPINSKI(A2)` column ➔ **Home ➔ Conditional Formatting ➔ Highlight Cells Rules ➔ Equal To...** ➔ type `PASS` and choose **Green Fill with Dark Green Text**.
2. **Similarity Gauge Bars (`0.0–1.0`):** Select your `=SMILES.SIMILARITY(A2, $B$1)` column ➔ **Home ➔ Conditional Formatting ➔ Data Bars** to render visual progress bars inside every cell.

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📐 Excel Formula Cheat Sheet</div>
    <p class="s2i-card-desc">Browse all 25+ chemical functions, return data types, and error code specifications</p>
  </a>
  <a href="../tools/formula-builder" class="s2i-card">
    <div class="s2i-card-title">🛠️ Formula Generator & Style Simulator</div>
    <p class="s2i-card-desc">Click to combine styling/highlight options interactively and copy the formula</p>
  </a>
</div>


