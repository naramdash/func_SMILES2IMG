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

::: info 🧪 LC-MS / HRMS Tip: Calculating Adduct $m/z$ Values
High-resolution mass spectrometry measures monoisotopic mass (`=SMILES2EXACTMASS(A2)`) rather than average molecular weight (`MW`). Combine it with standard adduct masses directly in Excel:
* **$[\text{M}+\text{H}]^+$ (Protonated):** `=SMILES2EXACTMASS(A2) + 1.0073`
* **$[\text{M}+\text{Na}]^+$ (Sodiated):** `=SMILES2EXACTMASS(A2) + 22.9892`
* **$[\text{M}-\text{H}]^-$ (Deprotonated):** `=SMILES2EXACTMASS(A2) - 1.0073`
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
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | **$\le 500$** g/mol | `180.16` | Size limit for passive membrane permeability |
| **Lipophilicity (LogP)** | `=SMILES2LOGP(A2)` | **$\le 5.0$** | `1.31` | Wildman-Crippen octanol/water partition balance |
| **H-Bond Donors (HBD)** | `=SMILES2HBD(A2)` | **$\le 5$** | `1` | Total count of $-\text{OH}$ and $-\text{NH}$ hydrogens |
| **H-Bond Acceptors (HBA)** | `=SMILES2HBA(A2)` | **$\le 10$** | `4` | Total count of nitrogen ($\text{N}$) and oxygen ($\text{O}$) atoms |

### Additional ADME Descriptors: Polar Surface Area (`TPSA`) & Rotatable Bonds (`ROTBONDS`)
| Descriptor | Formula | Example | Medicinal Chemistry Rule of Thumb |
| :--- | :--- | :---: | :--- |
| **Topological Polar Surface Area** | `=SMILES2TPSA(A2)` | `63.60` (Å²) | • **$\le 140\text{ \AA}^2$:** Good oral intestinal absorption<br>• **$\le 90\text{ \AA}^2$:** Favorable for CNS Blood-Brain Barrier (BBB) penetration |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `3` | • **$\le 10$ bonds (Veber's Rule):** Favorable conformational flexibility for oral bioavailability |

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
| **Carboxylic Acid ($-\text{COOH}$)** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **Ester ($-\text{COO}-$)** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **Amide ($-\text{CONH}-$)** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **Sulfonamide ($-\text{SO}_2\text{NH}-$)** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **Ketone ($-\text{C(=O)}-$)** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **Aldehyde ($-\text{CHO}$)** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **Phenol (Aromatic $-\text{OH}$)** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **Nitro Group ($-\text{NO}_2$)** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **Any Halogen ($\text{F, Cl, Br, I}$)** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **Trifluoromethyl ($-\text{CF}_3$)** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

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

## 5. Import & Export `.sdf` Chemical Database Files

Open and save multi-compound **`.sdf` (Structure-Data File)** libraries directly in Excel without external conversion scripts.

### 📥 Import an `.sdf` File into a Formatted Excel Table
1. Click **`[SMILES] ➔ [SDF File] ➔ [Import .sdf File...]`** on the Excel ribbon.
2. Select any `.sdf` file on your PC. A new worksheet table is created automatically with:
   * **Structure Column:** Live `=SMILES2IMG(...)` formulas that scale with cell size
   * **SMILES, Formula & MW Columns:** Canonical SMILES plus live `=SMILES2FORMULA` and `=SMILES2MW` formulas
   * **All SD Property Tags:** Every `<TAG>` in the file (e.g., `<Compound_ID>`, `<IC50_nM>`, `<Target>`, `<Purity>`) is extracted into its own dedicated Excel column!

### 📤 Export an Excel Table Back to `.sdf`
1. Select the table rows (or SMILES column + adjacent property columns) you want to export.
2. Click **`[SMILES] ➔ [SDF File] ➔ [Export .sdf File...]`** and choose a destination file name.
3. The add-in generates clean 2D coordinates (V2000 Molblocks) and bundles all adjacent column headers as SD `<TAG>` fields—ready for molecular docking or sharing with collaborators.
