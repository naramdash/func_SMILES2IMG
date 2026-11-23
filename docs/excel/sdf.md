# SDF File Import (`.sdf` / `.mol`)

Bring external **`.sdf` and `.mol` chemical databases** exported from RDKit, KNIME, DataWarrior, ChemDraw, or PubChem straight into a formatted **11-column live Excel worksheet** in one click.

<FeatureMedia id="excel-sdf-workflow" />

::: tip Test Immediately with the Sample `.sdf` Database
Download <a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf"><strong>sample_compounds.sdf</strong></a> and open it via **`[SMILES]` ➔ `[Import SDF]`**.
:::

## 1. Import `.sdf` / `.mol` into an 11-Column Live Worksheet

1. Click **`[SMILES]` ➔ `[Import SDF]`** on the Excel ribbon.
2. Select any single-molecule `.mol` or multi-compound `.sdf` file (imports up to 500 molecules cleanly).
3. `func_SMILES2IMG` automatically builds a formatted worksheet with auto-sized rows (`80 pt`) and **11 live columns**:
   - `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `TPSA`, `Rings`, `Aromatic Rings`, `Fsp3`

Every imported row uses standard Excel cell formulas (`=SMILES2IMG`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2TPSA`, etc.), so double-clicking structures or sorting/filtering rows works seamlessly right away.

## Related Guides

- [**Substructure Search**](/excel/sar) — Screen your imported `.sdf` library with `=SMILES.MATCH` and `hl:` highlights
- [**Identifier Conversion**](/excel/data-sdf) — Deduplicate merged `.sdf` libraries using `=SMILES2INCHIKEY`
- [**Batch Style, TSV, 3D**](/excel/tools-export) — Restyle an imported `.sdf` sheet to ACS B&W or Dark Mode in one click without losing highlights
