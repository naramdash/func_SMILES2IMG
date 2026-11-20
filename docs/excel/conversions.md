# Notation Conversion

Standardize compound lists across chemical names, CAS Registry Numbers, canonical SMILES, IUPAC InChI, and 27-character InChIKeys for `XLOOKUP` / `VLOOKUP` deduplication.

```text
=NAME2SMILES("50-78-2")
=SMILES2INCHI(B2)
=SMILES2INCHIKEY(B2)
=INCHI2IMG(C2)
```

<FeatureMedia id="excel-conversions" />

## Conversion Formulas

| Conversion | Formula | Example Output |
| :--- | :--- | :--- |
| **Name or CAS # ➔ SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES ➔ Substance Name** | `=SMILES2NAME(A2)` | `Aspirin` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` / `=KET2INCHI(A2)` | `InChI=1S/C9H8O4/c1-6(10)...` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` / `=INCHI2INCHIKEY(A2)` / `=KET2INCHIKEY(A2)` | `BSYNRYMUTXBXSQ-UHFFFAOYSA-N` |

::: tip Deduplicating Across Vendor Lists
Because different databases write different non-canonical SMILES for the same molecule, generate `=SMILES2INCHIKEY(A2)` in a helper column and use Excel's `XLOOKUP` or **Remove Duplicates** on the InChIKey column.
:::
