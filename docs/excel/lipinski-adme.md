# Lipinski, ADME, 3D Descriptors

Evaluate oral bioavailability (`Lipinski Rule of 5`), blood-brain barrier permeability (`TPSA`), topological scaffold metrics, and 3D medicinal chemistry complexity (`Fsp3`, chiral stereocenters) across hundreds of compounds in **0ms pure offline** inside Excel.

<FormulaTable
  title="Lipinski, ADME, 3D Complexity Sheet"
  :columns='["SMILES", "Formula"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LIPINSKI(A2)", desc: "Lipinski Rule of 5 verdict", out: "PASS" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LOGP(A3)", desc: "Wildman-Crippen lipophilicity (LogP)", out: "1.31" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2TPSA(A4)", desc: "Topological Polar Surface Area (Å²)", out: "63.60" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FSP3(A5)", desc: "Fraction of sp³ hybridized carbons (3D)", out: "0.11" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. Complete ADME, Lipinski & 3D Descriptor Catalog (`11` Functions)

All 11 descriptor functions automatically accept SMILES, InChI, or single-molecule KET cells and return native Excel `Integer`, `Double`, or `"PASS"`/`"FAIL"` values:

| Category | Excel Formula | Return Type | Aspirin (`A2`) | MedChem / Rule Guideline |
| :--- | :--- | :---: | :---: | :--- |
| **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `String` | `PASS` | `"PASS"` if `MW ≤ 500`, `LogP ≤ 5.0`, `HBD ≤ 5`, `HBA ≤ 10` |
| **LogP (Wildman-Crippen)** | `=SMILES2LOGP(A2)` | `Double` | `1.31` | **≤ 5.0** (Lipinski lipophilicity ceiling) |
| **Polar Surface Area (TPSA)** | `=SMILES2TPSA(A2)` | `Double` | `63.60` | **≤ 140 Å²** oral absorption; **≤ 90 Å²** CNS BBB penetration |
| **H-Bond Donors (HBD)** | `=SMILES2HBD(A2)` | `Integer` | `1` | **≤ 5** (`-OH` and `-NH` count) |
| **H-Bond Acceptors (HBA)** | `=SMILES2HBA(A2)` | `Integer` | `4` | **≤ 10** (nitrogen `N` + oxygen `O` count) |
| **Rotatable Bonds** | `=SMILES2ROTBONDS(A2)` | `Integer` | `3` | **≤ 10** (Veber oral flexibility limit) |
| **Heavy Atom Count** | `=SMILES2HEAVY(A2)` | `Integer` | `13` | Non-hydrogen atoms (`countHeavyAtoms`) |
| **Total Rings (SSSR)** | `=SMILES2RINGS(A2)` | `Integer` | `1` | Smallest Set of Smallest Rings |
| **Aromatic Rings** | `=SMILES2AROMATIC(A2)` | `Integer` | `1` | **≤ 3** recommended to avoid solubility drop |
| **Fraction Csp³ (Fsp³)** | `=SMILES2FSP3(A2)` | `Double` | `0.11` | **≥ 0.25** ("Escape from Flatland" 3D saturation score `0.00–1.00`) |
| **Chiral Centers** | `=SMILES2CHIRAL(A2)` | `Integer` | `0` | Count of defined tetrahedral stereocenters (`@` / `@@`) |

## 2. Copy-Ready MedChem Triage Rules (`IF` / `AND` / `FILTER`)

Because every descriptor returns native Excel scalars, you can automate Hit-to-Lead flagging directly with standard Excel logic:

| Medicinal Chemistry Filter Rule | Copy-Ready Excel Formula (`A2` = SMILES/InChI/KET) |
| :--- | :--- |
| **Veber Oral Bioavailability (`TPSA ≤ 140`, `RotB ≤ 10`)** | `=IF(AND(SMILES2TPSA(A2)<=140, SMILES2ROTBONDS(A2)<=10), "Veber PASS", "Low Oral")` |
| **CNS BBB Candidate (`TPSA ≤ 90`, `MW ≤ 450`, `Lipinski = PASS`)** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "CNS Lead", "Standard")` |
| **3D Escape-from-Flatland (`Fsp3 ≥ 0.25`, `Aromatic ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "Good 3D Profile", "High Flatness")` |
| **Combined Lipinski + Pharmacophore Hit** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "Priority Lead", "Exclude")` |

## Related Guides

- [**Formula, MW, Exact Mass**](/excel/properties) — Compute molecular weight, stoichiometry `mg`, and HRMS exact *m/z*
- [**Substructure & Similarity**](/excel/sar) — Combine ADME triage with `=SMILES.MATCH` scaffold filtering and Tanimoto ranking
- [**Playbook (.xlsx)**](/excel/playbook) — Explore live Hit-to-Lead scoring dashboards in tab `2. Properties & Lipinski`
