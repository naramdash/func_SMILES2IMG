# SDF Import & Export

Open multi-compound `.sdf` / `.mol` chemical databases directly into a formatted 12-column live-formula Excel sheet, or export worksheet compounds back to standard V2000 `.sdf`.

<FeatureMedia id="excel-sdf-workflow" />

::: tip Try It with the Sample `.sdf` Database
Download <a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf"><strong>sample_compounds.sdf</strong></a> and import it via **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** in Excel.
:::

## 1. Import an `.sdf` File into Excel

1. Click **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** on the ribbon.
2. Select any `.sdf` or `.mol` file.
3. A new worksheet is generated automatically with adjusted row heights, column widths, and **12 live formula columns**:
   - `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski`

## 2. Export a Worksheet Back to `.sdf`

1. Open any sheet containing compound structures or SMILES strings.
2. Click **`[SMILES]` ➔ `[SDF File]` ➔ `[Export Sheet to SDF (.sdf)...]`**.
3. Choose an output filename to save a standard **V2000 `.sdf`** file compatible with RDKit, KNIME, DataWarrior, and ChemDraw.
