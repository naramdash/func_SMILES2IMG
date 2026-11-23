# LogP, TPSA & 3D Descriptors

Evaluate octanol-water lipophilicity (`LogP`), polar surface area (`TPSA`), heavy atom/ring scaffold counts, and 3D carbon saturation (`Fsp3`) across compounds **locally offline** inside Excel—with full **SMILES, InChI, and KET** input parity.

<FormulaTable
  title="LogP, TPSA & 3D Complexity Sheet"
  :columns='["SMILES", "Formula"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LOGP(A2)", desc: "Wildman-Crippen lipophilicity (LogP)", out: "1.31" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2TPSA(A3)", desc: "Topological Polar Surface Area (Å²)", out: "63.60" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2AROMATIC(A4)", desc: "Aromatic ring count", out: "1" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FSP3(A5)", desc: "Fraction of sp³ hybridized carbons (3D)", out: "0.11" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. Complete LogP, TPSA & 3D Descriptor Catalog (`6` Metrics × `3` Formats = `18` Functions)

Every descriptor family provides strict input-validated **`SMILES2*`**, **`INCHI2*`**, and **`KET2*`** functions returning native Excel `Integer` or `Double` values (`SMILES2*` also auto-detects InChI and single-molecule KET cells; `INCHI2*` requires InChI and `KET2*` requires KET JSON):

| Category | SMILES Formula | InChI / KET Parity | Return Type | Aspirin (`A2`) | Reference / Guideline |
| :--- | :--- | :--- | :---: | :---: | :--- |
| **LogP (Wildman-Crippen)** | `=SMILES2LOGP(A2)` | `=INCHI2LOGP(A2)` · `=KET2LOGP(A2)` | `Double` | `1.31` | Wildman-Crippen octanol-water partition coefficient (`LogP ≤ 5`) |
| **Polar Surface Area (TPSA)** | `=SMILES2TPSA(A2, [include_sp])` | `=INCHI2TPSA(A2, [include_sp])` · `=KET2TPSA(A2, [include_sp])` | `Double` | `63.60` | `≤ 140 Å²` oral; `≤ 90 Å²` CNS BBB (`TRUE` includes S/P per PubChem) |
| **Heavy Atom Count** | `=SMILES2HEAVY(A2)` | `=INCHI2HEAVY(A2)` · `=KET2HEAVY(A2)` | `Integer` | `13` | Non-hydrogen atoms (`countHeavyAtoms`) |
| **Total Rings (SSSR)** | `=SMILES2RINGS(A2)` | `=INCHI2RINGS(A2)` · `=KET2RINGS(A2)` | `Integer` | `1` | Smallest Set of Smallest Rings |
| **Aromatic Rings** | `=SMILES2AROMATIC(A2)` | `=INCHI2AROMATIC(A2)` · `=KET2AROMATIC(A2)` | `Integer` | `1` | **≤ 3** recommended to avoid solubility drop *(EPAM Indigo generic model)* |
| **Fraction Csp³ (Fsp³)** | `=SMILES2FSP3(A2)` | `=INCHI2FSP3(A2)` · `=KET2FSP3(A2)` | `Double` | `0.11` | **≥ 0.25** ("Escape from Flatland" `sp³ C / total C`; `0.00` for carbon-free species) |

::: info Scientific Accuracy Notes (EPAM Indigo Engine)
* **TPSA Modes & Accepted Flags:** Default `=SMILES2TPSA(A2)` uses Ertl N/O-only PSA (matching RDKit). Pass `TRUE` (or `1`, `"yes"`, `"sp"`, `"s"`, `"p"`, `"pubchem"`, `"all"`) as the 2nd argument (`=SMILES2TPSA(A2, TRUE)`) to include sulfur and phosphorus atoms (matching PubChem).
* **Formal Charge, Tautomers & Salt Scope:** Wildman-Crippen `LogP` and Ertl `TPSA` are sensitive to protonation/charge representation (e.g., zwitterionic vs neutral amino acids, `[N+](=O)[O-]` vs `N(=O)=O`, `[S+]([O-])` vs `S=O`) and mobile-H `InChI=1S` amide/enol tautomer shifts (`SMILES`/`KET` preserve your exact drawn tautomer). Salts and organometallic complexes are evaluated across all drawn fragments (`LogP` uses Indigo's Crippen fallback parameters, so treat inorganic/metal `LogP` as a heuristic descriptor rather than experimental partition data).
* **No Black-Box HBD / HBA / Lipinski / Chiral Wrappers:** Definitions for HBD/HBA/RotBonds vary across toolkits, and perceived vs drawn stereocenters on flat SMILES can mislead triage. Combine `SMILES2MW`, `SMILES2LOGP`, `SMILES2TPSA`, `SMILES2FSP3`, and `SMILES2AROMATIC` explicitly in Excel formulas so your screening criteria stay transparent.
* **Query & Reaction Guardrails:** Passing query R-groups (`*`) or reactions (`>>`) returns `#N/A` on all scalar descriptor functions.
:::

## 2. Transparent MedChem Triage Rules (`IF` / `AND` / `FILTER`)

Because every descriptor returns native Excel scalars, build auditable Hit-to-Lead criteria directly in standard Excel formulas:

| Medicinal Chemistry Filter Rule | Copy-Ready Excel Formula (`A2` = SMILES/InChI/KET) |
| :--- | :--- |
| **Oral Physicochemical Window (`MW ≤ 500`, `LogP ≤ 5`, `TPSA ≤ 140`)** | `=IF(AND(SMILES2MW(A2)<=500, SMILES2LOGP(A2)<=5, SMILES2TPSA(A2)<=140), "PASS", "FAIL")` |
| **CNS BBB Candidate (`TPSA ≤ 90`, `MW ≤ 450`, `LogP ≤ 4`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450, SMILES2LOGP(A2)<=4), "CNS Lead", "Standard")` |
| **3D Escape-from-Flatland (`Fsp3 ≥ 0.25`, `Aromatic ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "Good 3D Profile", "High Flatness")` |
| **Combined Property + Pharmacophore Hit** | `=IF(AND(SMILES2MW(A2)<=500, SMILES2TPSA(A2)<=120, SMILES.MATCH(A2, "c1ccncc1")), "Priority Lead", "Exclude")` |

## Related Guides

- [**Formula, MW, Exact Mass**](/excel/properties) — Compute molecular weight, stoichiometry `mg`, and HRMS exact *m/z*
- [**Substructure Search**](/excel/sar) — Combine physicochemical triage with `=SMILES.MATCH` / `=INCHI.MATCH` / `=KET.MATCH`
- [**Playbook (.xlsx)**](/excel/playbook) — Explore live Hit-to-Lead scoring dashboards in tab `2. Properties & 3D`
