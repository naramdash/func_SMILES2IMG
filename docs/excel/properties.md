# Formula, MW, Exact Mass

Compute standard Hill chemical formulas, average molecular weights (`MW`), HRMS monoisotopic exact masses, and substance names locally—and plug native numeric outputs straight into organic synthesis ELN stoichiometry calculators and LC-MS / HRMS adduct tables.

<FormulaTable
  title="Formula, Mass, Name Sheet (Offline)"
  :columns='["SMILES", "Formula"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FORMULA(A2)", desc: "Standard Hill molecular formula", out: "C9H8O4" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A3)", desc: "Average molecular weight (g/mol, numeric)", out: "180.16" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2EXACTMASS(A4)", desc: "Monoisotopic HRMS exact mass (4 decimals)", out: "180.0423" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2NAME(A5)", desc: "IUPAC / PubChem substance name (cached)", out: "Aspirin" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. Core Identity & Mass Functions (`12` Total Across SMILES / InChI / KET)

All four descriptors work with SMILES, InChI, and single-molecule KET cells via a single `(input)` argument (`=INCHI2*` and `=KET2*` delegate directly to the same parser):

| Property Output | SMILES Input | InChI / KET Input | Aspirin (`A2`) | Execution & Output Type |
| :--- | :--- | :--- | :---: | :--- |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)` · `=KET2FORMULA(A2)` | `C9H8O4` | **Offline** · Standard Hill notation (`String`) |
| **Molecular Weight (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)` · `=KET2MW(A2)` | `180.16` | **Offline** · Numeric `Double` in g/mol (2 decimals) |
| **Exact Mass (HRMS)** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)` · `=KET2EXACTMASS(A2)` | `180.0423` | **Offline** · Monoisotopic mass in Da (4 decimals) |
| **Substance Name** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `Aspirin` | Built-in dictionary + PubChem lookup (30-day local disk cache, `String`) |

## 2. Template 1: Organic Synthesis Stoichiometry Calculator (ELN)

Because `=SMILES2MW` returns a genuine numeric `Double`, you can multiply it directly by a `mmol` cell to build an automated reagent weighing table (`mass = MW × mmol`):

| Col A (Reagent) | Col B (`=NAME2SMILES(A2)`) | Col C (`=SMILES2IMG(B2)`) | Col D (`=SMILES2MW(B2)`) | Col E (Eq.) | Col F (`mmol`) | Col G (`=D2*F2` mg) |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D structure)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D structure)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |

To calculate theoretical yield (`mg`) of your product row (column `B4`), multiply the limiting reagent's `mmol` (`$F$2`), equivalents, and `=SMILES2MW(B4)` directly.

## 3. Template 2: LC-MS & HRMS Calculated Ion Mass (*m/z*)

Use `=SMILES2EXACTMASS(A2)` (accurate to 4 decimal places) to compute calculated exact *m/z* values for common mass spectrometry adducts automatically:

| MS Ion / Adduct | Copy-Ready Excel Formula (`A2` = Structure) | Aspirin (`C9H8O4`) |
| :--- | :--- | :---: |
| **Neutral Exact Mass (`M`)** | `=SMILES2EXACTMASS(A2)` | `180.0423` |
| **`[M + H]+`** *(ESI Positive)* | `=ROUND(SMILES2EXACTMASS(A2) + 1.007276, 4)` | `181.0496` |
| **`[M + NH4]+`** | `=ROUND(SMILES2EXACTMASS(A2) + 18.033823, 4)` | `198.0761` |
| **`[M + Na]+`** | `=ROUND(SMILES2EXACTMASS(A2) + 22.989218, 4)` | `203.0315` |
| **`[M - H]-`** *(ESI Negative)* | `=ROUND(SMILES2EXACTMASS(A2) - 1.007276, 4)` | `179.0350` |

::: info Salts, Isotopes & Non-Molecule Inputs
* **Salts & Counterions:** Multi-component structures (`CN(C)C(=N)NC(=N)N.Cl`, `[Na+].CC(=O)[O-]`) are evaluated across the complete salt as drawn, reflecting counterion mass and formula directly.
* **Isotopic Labels:** Heavy isotopes (`[2H]`, `[13CH3]`) update `MW` and `Exact Mass` with exact isotopic masses, while `=SMILES2FORMULA` outputs standard IUPAC Hill notation (`[2H]C([2H])([2H])S(=O)C([2H])([2H])[2H]` ➔ `C2H6OS`).
* **Queries & Reactions:** Query R-groups (`*`, `[R1]`) and multi-component reactions (`>>`) return `#N/A` on single-molecule scalar formulas so unphysical numbers are never produced.
:::

## Related Guides

- [**LogP, TPSA & 3D**](/excel/lipinski-adme) — Compute Crippen `LogP`, Ertl `TPSA`, rings, and `Fsp3`
- [**Identifier Conversion**](/excel/data-sdf) — Convert between Name/CAS, Canonical SMILES, InChI, and `InChIKey`
- [**Playbook (.xlsx)**](/excel/playbook) — Open tabs `2. Properties & 3D` and `7. Real-World Workflows` to copy these live tables directly
