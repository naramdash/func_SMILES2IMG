# SDF Import & Export (`.sdf` / `.mol`)

Bring external **`.sdf` and `.mol` chemical databases** exported from RDKit, KNIME, DataWarrior, ChemDraw, or PubChem straight into a formatted **12-column live Excel worksheet**, or export your worksheet back to a standard V2000 `.sdf` file in one click.

<FeatureMedia id="excel-sdf-workflow" />

::: tip Test Immediately with the Sample `.sdf` Database
Download <a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf"><strong>sample_compounds.sdf</strong></a> and open it via **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`**.
:::

## 1. Import `.sdf` / `.mol` into a 12-Column Live Worksheet

1. Click **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** on the Excel ribbon.
2. Select any single-molecule `.mol` or multi-compound `.sdf` file.
3. `func_SMILES2IMG` automatically builds a formatted worksheet with auto-sized rows (`62–80 pt`) and **12 live formula columns**:
   - `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski`

Every imported row uses standard cell formulas (`=SMILES2IMG`, `=SMILES2MW`, `=SMILES2LIPINSKI`, etc.), so double-clicking structures or sorting/filtering rows works seamlessly right away.

## 2. Export Any Excel Sheet Back to V2000 `.sdf`

1. Activate any worksheet containing a `SMILES` column or rendered structure cells.
2. Click **`[SMILES]` ➔ `[SDF File]` ➔ `[Export Sheet to SDF (.sdf)...]`** on the ribbon and choose an output path.
3. `func_SMILES2IMG` writes a standard **V2000 `.sdf`** file containing 2D atomic coordinates and your sheet's metadata property tags—ready to open directly in RDKit, KNIME, DataWarrior, or ChemDraw.

## Related Guides

- [**Substructure & Similarity**](/excel/sar) — Screen your imported `.sdf` library with `=SMILES.MATCH` and `=SMILES.SIMILARITY`
- [**Identifier Conversion**](/excel/data-sdf) — Deduplicate merged `.sdf` libraries using `=SMILES2INCHIKEY`
- [**Batch Style, TSV, 3D**](/excel/tools-export) — Restyle an imported `.sdf` sheet to ACS B&W or Dark Mode in one click
