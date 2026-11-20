# Compound Tables & Properties

Build manuscript compound summary tables in Microsoft Word and populate chemical formulas, molecular weights, LogP, and Lipinski metrics without manual typing.

<FeatureMedia id="word-ch3-compound-table" />

## Steps

1. Place your cursor inside a Word table cell and insert each compound via **`[Insert / Edit Structure]`** or **`[Paste Structure]`**.
2. Select any inserted structure and click **`[SMILES]` ➔ `[Molecule Info]`** (or right-click ➔ **`Chemical Structure` ➔ `Copy Molecule Info`**).
3. Choose the copy format that matches your table:
   - **`Copy All Summary (Text)`:** Copies a 2-row table (Headers + Data) for starting a new table.
   - **`Copy as Excel Row (TSV Columns)`:** Copies 17 tab-separated descriptor values (no header) to append across table columns.
   - **`Detailed Molecule Info Window...`:** Inspect all 17 descriptors with a live 2D preview, double-click any row to copy a single value, or export Markdown / JSON.

::: tip Preserve Column Widths on Re-Edit
Right-click your table ➔ **AutoFit** ➔ **Fixed Column Width** so Word keeps adjacent text columns steady when you double-click and modify structures later.
:::
