# Structure Tables & Name/CAS

Turn columns of SMILES strings, common chemical names, lab solvents, or CAS Registry Numbers into scaled 2D structures directly inside Excel cells.

## 1. SMILES Column to Structure Table (`=SMILES2IMG`)

Keep compound names in Column A and SMILES strings in Column B, then render crisp, aspect-ratio-locked 2D structures in Column C:

```text
=SMILES2IMG(B2)
```

<FeatureMedia id="excel-formula-table" />

1. Enter compound names in `A2:A5` and SMILES strings in `B2:B5` (e.g. `CC(=O)Oc1ccccc1C(=O)O` for Aspirin).
2. Type `=SMILES2IMG(B2)` in `C2` and press <kbd>Enter</kbd>.
3. Increase **row height (`62–80 pt`)** and **column width**—structures scale cleanly without distortion.
4. Drag the fill handle down to render hundreds of rows automatically.

## 2. Chemical Names, Solvents & CAS Numbers (`=NAME2IMG`)

Skip looking up SMILES manually by passing common names, lab solvent abbreviations (`THF`, `DMSO`, `DCM`), or CAS numbers directly:

```text
=NAME2IMG("Aspirin")
=NAME2IMG("50-78-2")
=NAME2SMILES("Ibuprofen")
```

<FeatureMedia id="excel-name-cas" />

::: details Built-In 0ms Offline Dictionary (Solvents, Drugs, Biomolecules & CAS)
Common lab solvents, drugs, biomolecules, and CAS numbers resolve **locally in 0ms**. Uncached compounds are looked up once via PubChem and cached locally on disk for future offline use:
* **Lab Solvents:** `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`
* **Pharmaceuticals:** `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel`
* **Biomolecules:** `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`
:::

::: details ELN Stoichiometry Example Using `=SMILES2MW`
Because `=SMILES2MW` returns a true Excel number (`double`), multiply it directly by `mmol` to calculate reagent mass (`mg`):

| Col A (Name) | Col B (`=NAME2SMILES(A2)`) | Col C (`=SMILES2IMG(B2)`) | Col D (`=SMILES2MW(B2)`) | Col E (Equiv) | Col F (`mmol`) | Col G (`=D2*F2` mg) |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D structure)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D structure)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |
:::
