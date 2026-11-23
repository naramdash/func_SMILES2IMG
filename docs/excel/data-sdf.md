# Identifier Conversions & `InChIKey` Deduplication

Standardize and deduplicate chemical identifiers across **Name**, **CAS Registry Number**, **Canonical SMILES**, **IUPAC InChI**, and **27-character `InChIKey` hashes**—enabling exact `XLOOKUP` joins and one-click duplicate removal across merged vendor or literature tables.

<FormulaTable
  title="Identifier Conversion Sheet"
  :columns='["Input (SMILES / InChI)", "Formula"]'
  :rows='[
    { input: "", formula: "=NAME2SMILES(\"50-78-2\")", desc: "Standalone CAS # or Name to canonical SMILES", out: "CC(=O)Oc1ccccc1C(=O)O" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2INCHI(A3)", desc: "SMILES or KET cell to IUPAC InChI", out: "InChI=1S/..." },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2INCHIKEY(A4)", desc: "Canonical 27-char InChIKey hash", out: "BSYNRYMUTXBXSQ-..." },
    { input: "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3", formula: "=INCHI2SMILES(A5)", desc: "IUPAC InChI back to canonical SMILES", out: "CCO" }
  ]'
/>

<FeatureMedia id="excel-conversions" />

## 1. Cross-Notation Conversion Functions

Every identifier conversion returns a native Excel text `String` (`*2NAME` resolves via offline catalog + cached PubChem lookup; SMILES/InChI/KET structural conversions compute 100% offline in 0ms):

| Conversion Direction | Excel Formula | Example Output (`Aspirin`) |
| :--- | :--- | :--- |
| **Name or CAS # ➔ Canonical SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ Name** | `=SMILES2NAME(A2)` · `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `Aspirin` |
| **SMILES / KET ➔ IUPAC InChI** | `=SMILES2INCHI(A2)` · `=KET2INCHI(A2)` | `InChI=1S/C9H8O4/c1-6(10)...` |
| **InChI ➔ Canonical SMILES** | `=INCHI2SMILES(A2)` | `CC(=O)Oc1ccccc1C(=O)O` |
| **KET ➔ Canonical SMILES** | `=KET2SMILES(A2)` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` · `=INCHI2INCHIKEY(A2)` · `=KET2INCHIKEY(A2)` | `BSYNRYMUTXBXSQ-UHFFFAOYSA-N` |

## 2. Deduplicating Multi-Vendor Tables with `InChIKey`

Different chemical suppliers and papers often serialize atom order differently in SMILES strings for the exact same molecule (e.g., `c1ccccc1C(=O)O` vs. `O=C(O)c1ccccc1`), making plain text matching miss duplicates.

Add a canonical helper column using `=SMILES2INCHIKEY(B2)`—because the 27-character `InChIKey` is an invariant hash independent of SMILES atom numbering:

| Task | Copy-Ready Excel Recipe (`D` = `=SMILES2INCHIKEY(B2)`) | Result |
| :--- | :--- | :--- |
| **Flag Duplicates Live** | `=IF(COUNTIF($D$2:$D$500, D2)>1, "Duplicate", "Unique")` | Highlights overlapping compounds instantly across sheets |
| **Exact Vendor Cross-Lookup** | `=XLOOKUP(D2, VendorB!$D:$D, VendorB!$E:$E, "Not Found")` | Joins assay data or pricing reliably even when SMILES strings differ |
| **Clean Master Table** | Select your table ➔ Excel **`[Data]` ➔ `[Remove Duplicates]`** on the `InChIKey` column | Strips chemically identical duplicates with 100% accuracy |

## Related Guides

- [**Substructure Search**](/excel/sar) — Screen scaffolds with `=SMILES.MATCH`, `=INCHI.MATCH`, and `=KET.MATCH` (`TRUE`/`FALSE`)
- [**SDF Import**](/excel/sdf) — Import external `.sdf`/`.mol` libraries into an 11-column live worksheet
- [**All 49 Functions**](/excel/formulas) — Full parameter catalog across all 49 Excel functions
