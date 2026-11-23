# Names, Solvents, CAS Images

Render 2D chemical structures directly from **English chemical names**, **common lab solvent acronyms** (`THF`, `DMF`, `DMSO`, `DCM`), **reagents**, or **CAS Registry Numbers** (`"50-78-2"`) via `=NAME2IMG`—combining a 0ms built-in offline dictionary with a 30-day local PubChem cache—without looking up a SMILES string first.

<FormulaTable
  title="Name, CAS, Solvent Sheet"
  :columns='["Name / CAS", "Formula"]'
  :rows='[
    { input: "Aspirin", formula: "=NAME2IMG(A2)", desc: "Resolve common English drug/reagent name", out: "2D Aspirin" },
    { input: "50-78-2", formula: "=NAME2IMG(A3)", desc: "Resolve CAS Registry Number cleanly", out: "2D Aspirin" },
    { input: "", formula: "=NAME2IMG(\"THF\", \"\", \"bw\")", desc: "Standalone solvent acronym in ACS B&W", out: "B&W THF" }
  ]'
/>

<FeatureMedia id="excel-name-cas" />

## 1. 0ms Built-In Offline Dictionary + Local PubChem Cache

`=NAME2IMG(name_or_cas, [background], [style], [transform])` (and `=NAME2IMG.FLOAT` on Excel 2016–2021) checks an internal 0ms offline chemical dictionary first. For any compound outside the built-in catalog, `func_SMILES2IMG` performs a one-time PubChem lookup and caches valid resolutions on disk for **30 days** (`%LOCALAPPDATA%\Smiles2Img`) so future workbooks open instantaneously offline. Numeric CAS strings are also checked mathematically (check digit) in **0ms**—invalid CAS numbers (such as `"50-78-3"`) immediately return `#N/A` without network delay.

| Category | Built-In 0ms Offline Entries (Examples) | Excel Formula Example |
| :--- | :--- | :--- |
| **Lab Solvents** | `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`, `Chloroform`, `Diethyl Ether`, `Hexane`, `Cyclohexane` | `=NAME2IMG("THF")` |
| **Extended Reagents *(PubChem ➔ 30-day cache)*** | `HATU`, `DIPEA`, `DMAP`, `EDCI`, `HOBt`, `mCPBA`, `TFA` *(cached locally 30 days after first lookup)* | `=NAME2IMG("HATU")` |
| **Pharmaceuticals** | `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel` | `=NAME2IMG("50-78-2")` |
| **Biomolecules & Amino Acids** | `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`, `Cholesterol` | `=NAME2IMG("Caffeine")` |

::: tip Works with Every Shared `[bg]` and `[style]` Keyword
`=NAME2IMG` passes `[background]`, `[style]`, and `[transform]` into the exact same renderer as `=SMILES2IMG`:
* Journal monochrome: `=NAME2IMG(A2, "", "bw")`
* Dark Mode white ink: `=NAME2IMG(A2, "#1E293B", "white")`
* Highlighting carboxylic acids: `=NAME2IMG(A2, "", "hl:C(=O)[OH]")`

See [**Styles & Highlights**](/excel/styles-reactions) for the complete token table.
:::

## 2. Pair `=NAME2IMG` with `=NAME2SMILES` & Stoichiometry Columns

In organic synthesis ELNs and reagent inventory sheets, type the chemical name or CAS number once in column A and auto-populate structure, SMILES, and MW together:

| Col A (Input Name / CAS) | Col B (`=NAME2SMILES(A2)`) | Col C (`=NAME2IMG(A2)`) | Col D (`=SMILES2MW(B2)`) |
| :--- | :--- | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D structure)* | `138.12` |
| `50-78-2` *(Aspirin)* | `CC(=O)Oc1ccccc1C(=O)O` | *(2D structure)* | `180.16` |
| `DIPEA` *(PubChem ➔ 30d cache)* | `CCN(C(C)C)C(C)C` | *(2D structure)* | `129.24` |

## Related Guides

- [**SMILES, InChI**](/excel/structures) — Unified 4-argument architecture and multi-row SMILES/InChI tables
- [**Styles & Highlights**](/excel/styles-reactions) — Customize colors, ACS B&W, SMARTS highlighting, and rotation
- [**Identifier Conversion**](/excel/data-sdf) — Convert Names & CAS # to canonical SMILES (`=NAME2SMILES`) and `InChIKey` (`=SMILES2INCHIKEY`)
