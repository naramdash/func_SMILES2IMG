# Properties & Lipinski Ro5

Compute physicochemical descriptors, HRMS monoisotopic exact mass, Lipinski Rule of 5 compliance, and 3D complexity metrics locally in Excel with zero latency.

```text
=SMILES2MW(A2)
=SMILES2EXACTMASS(A2)
=SMILES2LIPINSKI(A2)
=SMILES2FSP3(A2)
```

<FeatureMedia id="excel-cheminformatics" />

## Physicochemical & Drug-Likeness Formulas

| Descriptor | Formula | MedChem Rule | Aspirin (`A2`) | Notes |
| :--- | :--- | :---: | :---: | :--- |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | — | `C9H8O4` | Standard Hill notation |
| **Molecular Weight** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | Numeric `double` for stoichiometry |
| **Exact Mass (HRMS)** | `=SMILES2EXACTMASS(A2)` | — | `180.0423` | `[M+H]+` = `+1.0073`, `[M+Na]+` = `+22.9892` |
| **LogP / TPSA** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | **≤ 5.0** / **≤ 140 Å²** | `1.31` / `63.60` | CNS BBB ≤ 90 Å² |
| **HBD / HBA / RotB** | `=SMILES2HBD` / `HBA` / `ROTBONDS` | **≤ 5** / **≤ 10** / **≤ 10** | `1` / `4` / `3` | Oral bioavailability descriptors |
| **Lipinski Pass/Fail** | `=SMILES2LIPINSKI(A2)` | — | `PASS` | Combined 4-rule evaluation |
| **Heavy / Rings / Arom** | `=SMILES2HEAVY` / `RINGS` / `AROMATIC` | **Arom ≤ 3** | `13` / `1` / `1` | Developability & scaffold tracking |
| **3D Fsp³ / Chiral** | `=SMILES2FSP3(A2)` / `=SMILES2CHIRAL(A2)` | **Fsp³ ≥ 0.25** | `0.11` / `0` | "Escape from Flatland" metrics |

::: tip Using InChI or KET Input?
Replace `SMILES2` with `INCHI2` or `KET2` (e.g. `=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).
:::

## Combining with Excel `IF` / `AND` Rules

| Screening Rule | Copy-Paste Formula (`A2` = SMILES) |
| :--- | :--- |
| **CNS BBB Candidate (`TPSA ≤ 90`, `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "CNS Candidate", "Standard")` |
| **3D-Rich Lead (`Fsp3 ≥ 0.25`, `Arom ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "Escape Flatland", "Flat")` |
