# Excel Use Cases

Explore 6 practical ways **`func_SMILES2IMG`** turns Microsoft Excel into a self-contained chemistry workbench—with copy-ready multi-column sheets you can paste straight into cell `A1`.

| Use Case | What You Can Do | Key Tool |
| :--- | :--- | :--- |
| [**1. Synthesis ELN & Stoichiometry**](#synthesis-eln) | Auto-fill SMILES, structures, MW, and mg weights from names or CAS | `=NAME2SMILES`, `=SMILES2MW` |
| [**2. Hit-to-Lead SAR Triage**](#sar-triage) | Highlight query scaffolds and screen Lipinski + ADME thresholds live | `=SMILES.MATCH`, `"hl:"` |
| [**3. Reagent & Library Deduplication**](#deduplication) | Normalize vendor SMILES variations into `InChIKey` for `COUNTIF` & `XLOOKUP` | `=SMILES2INCHIKEY` |
| [**4. SDF Import & Office Handoff**](#sdf-handoff) | Convert `.sdf` files to 12-column live tables and reuse schemes in Word/PPT | `[Import SDF]`, `[Batch Style]` |
| [**5. R-Group Analog Enumeration**](#rgroup-enumeration) | Swap R-groups with `SUBSTITUTE` and preview analog structures & properties in 0ms | `=SUBSTITUTE`, `=SMILES2IMG` |
| [**6. Dynamic Lead Comparison & FILTER**](#dynamic-comparison) | Compare in-cell structure images dynamically via `XLOOKUP` and `FILTER` (365+) | `=SMILES.SIMILARITY` |

---

## 1. Synthesis ELN & Stoichiometry {#synthesis-eln}

Eliminate switching between ChemDraw, PubChem, and manual weight calculators during reaction setup. Sketch your reaction scheme once at the top via **`[SMILES]` ➔ `[Insert / Edit (KET)]`**, then type reagent names or CAS numbers in column `A` to auto-populate SMILES, structures, live numeric `MW`, and `mg` weighing values.

<FeatureMedia id="excel-reactions-ket" />

1. Paste the table below into **`A1`** (append **`.FLOAT`** to `=SMILES2IMG` on Excel 2016–2021).
2. Adjust the scale in `E2` (`5.0` mmol) to recalculate every reagent weight (`F2:F4`) immediately.
3. Click **`[SMILES]` ➔ `[Insert / Edit (KET)]`** on a header cell to attach a double-click-editable reaction scheme.

<FormulaTable
  title="ELN Stoichiometry Calculator Sheet"
  :columns='["Reagent / CAS", "SMILES", "Structure", "MW", "mmol", "Mass (mg)"]'
  :rows='[
    { cells: ["Salicylic Acid", "=NAME2SMILES(A2)", "=SMILES2IMG(B2)", "=SMILES2MW(B2)", "5.0", "=ROUND(D2*E2, 1)"], desc: "Limiting reagent (1.0 eq)", out: "690.6 mg" },
    { cells: ["Acetic Anhydride", "=NAME2SMILES(A3)", "=SMILES2IMG(B3)", "=SMILES2MW(B3)", "=E2*1.5", "=ROUND(D3*E3, 1)"], desc: "Excess reagent (1.5 eq)", out: "765.7 mg" },
    { cells: ["50-78-2", "=NAME2SMILES(A4)", "=SMILES2IMG(B4)", "=SMILES2MW(B4)", "=E2", "=ROUND(D4*E4, 1)"], desc: "Theoretical Aspirin product", out: "900.8 mg" }
  ]'
/>

* **Next:** [**Formula, MW, Exact Mass**](/excel/properties) · [**Reaction Schemes (KET)**](/excel/reactions) · Playbook **Tab 6 & Tab 7**

---

## 2. Hit-to-Lead SAR Triage {#sar-triage}

Point every analog row to a single query scaffold cell (`$B$2`) so editing one SMARTS pattern updates substructure flags, color-highlighted 2D structures, and Lipinski triage tiers across your entire screening table simultaneously.

<FeatureMedia id="excel-sar-screening" />

1. Paste the sheet below at **`A1`** (`$B$2` holds the editable query SMARTS `c1ccncc1`).
2. Edit `$B$2` (e.g., change to `c1ccccc1`) to watch columns `C:E` re-flag and re-highlight automatically.
3. Use Excel **Sort & Filter**—native in-cell structures stay locked to their rows on Microsoft 365 / Excel 2024+.

<FormulaTable
  title="Dynamic Query SAR Triage Sheet"
  :columns='["Target SMILES", "Query ($B$2)", "Match", "Highlighted 2D", "Priority Tier"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccncc1", "=SMILES.MATCH(A2, $B$2)", "=SMILES2IMG(A2, \"\", \"hl:\"&$B$2)", "=IF(AND(SMILES2LIPINSKI(A2)=\"PASS\", SMILES.MATCH(A2, $B$2)), \"Priority\", \"Monitor\")"], desc: "Aspirin vs. pyridine query", out: "Monitor" },
    { cells: ["CN1CCCC1c2cccnc2", "", "=SMILES.MATCH(A3, $B$2)", "=SMILES2IMG(A3, \"\", \"hl:\"&$B$2)", "=IF(AND(SMILES2LIPINSKI(A3)=\"PASS\", SMILES.MATCH(A3, $B$2)), \"Priority\", \"Monitor\")"], desc: "Nicotine matches $B$2 scaffold", out: "Priority" }
  ]'
/>

* **Next:** [**Substructure & Similarity (SAR)**](/excel/sar) · [**Lipinski, ADME, 3D**](/excel/lipinski-adme) · Playbook **Tab 2 & Tab 3**

---

## 3. Reagent & Library Deduplication {#deduplication}

Chemical vendors and papers often write the exact same molecule with different SMILES atom orderings (`c1ccccc1C(=O)O` vs. `O=C(O)c1ccccc1`). Generating the invariant 27-character **`InChIKey`** lets Excel detect duplicates via `COUNTIF`, merge tables via `XLOOKUP`, and clean master lists with **Remove Duplicates**.

<FeatureMedia id="excel-conversions" />

1. Paste the sheet below at **`A1`** to compare two different SMILES writings of Benzoic acid alongside Aspirin.
2. Expand `COUNTIF($C$2:$C$4, C2)>1` across all rows to flag chemically identical entries live.
3. Cross-join vendor pricing or assay data across sheets by matching the `InChIKey` column with `XLOOKUP`.

<FormulaTable
  title="InChIKey Deduplication Sheet"
  :columns='["Input SMILES / Name", "Canonical SMILES", "InChIKey", "Duplicate Flag"]'
  :rows='[
    { cells: ["c1ccccc1C(=O)O", "=INCHI2SMILES(SMILES2INCHI(A2))", "=SMILES2INCHIKEY(A2)", "=IF(COUNTIF($C$2:$C$4, C2)>1, \"Duplicate\", \"Unique\")"], desc: "Benzoic acid (ring-first SMILES)", out: "Duplicate" },
    { cells: ["O=C(O)c1ccccc1", "=INCHI2SMILES(SMILES2INCHI(A3))", "=SMILES2INCHIKEY(A3)", "=IF(COUNTIF($C$2:$C$4, C3)>1, \"Duplicate\", \"Unique\")"], desc: "Same molecule (carboxyl-first)", out: "Duplicate" },
    { cells: ["Aspirin", "=NAME2SMILES(A4)", "=SMILES2INCHIKEY(B4)", "=IF(COUNTIF($C$2:$C$4, C4)>1, \"Duplicate\", \"Unique\")"], desc: "Resolved from name", out: "Unique" }
  ]'
/>

* **Next:** [**Names, Solvents, CAS**](/excel/name-cas) · [**Identifier Conversion**](/excel/data-sdf) · Playbook **Tab 4**

---

## 4. SDF Import & Office Handoff {#sdf-handoff}

Import `.sdf` screening libraries exported from RDKit, KNIME, DataWarrior, or PubChem into a live formula sheet, triage hits, and transfer publication-ready structures directly to Word and PowerPoint without redrawing.

<FeatureMedia id="excel-sdf-workflow" />

1. **1-Click 12-Column Import:** Click **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** to generate formatted rows with `Name`, `Structure`, `SMILES`, `Formula`, `MW`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, and `Lipinski`.
2. **Restyle in 1 Click:** Filter down to top hits and run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`** (or `Dark Mode`).
3. **Hand Off Without Redrawing:**
   - **To Word:** Copy filtered table rows into Word, or click **`[Copy SMILES]`** / **`[Copy KET]`** and run **`[Paste Structure]`** ([Word Use Cases](/word/use-cases)).
   - **To PowerPoint:** Paste hits onto slides with **`[Paste Structure]`**, then apply **`[Batch Style]` ➔ `[Dark Mode]`** ([PowerPoint Use Cases](/powerpoint/use-cases)).

* **Next:** [**SDF Import & Export**](/excel/sdf) · [**Batch Style, TSV, 3D**](/excel/tools-export) · Playbook **Tab 9**

---

## 5. R-Group Analog Enumeration {#rgroup-enumeration}

Simulate entire analog series inside Excel using native **`SUBSTITUTE`**—no Python or RDKit scripts required. Adding a single substituent SMILES fragment renders the analog and updates MW, LogP, and Lipinski status in `0ms`.

<FeatureMedia id="excel-cheminformatics" />

1. Paste the enumeration table below at **`A1`** (`$B$2` holds the `[R]` scaffold template).
2. Add more SMILES fragments down column `A` (`Cl`, `OC`, `C#N`) and fill columns `C:F` downward.
3. Combine with `=SMILES2LOGP(C2)` or `=SMILES2TPSA(C2)` to spot ADME trade-offs before synthesis.

<FormulaTable
  title="R-Group Enumeration Sheet"
  :columns='["R-Group", "Template ($B$2)", "Analog SMILES", "2D Structure", "MW", "Lipinski"]'
  :rows='[
    { cells: ["O", "CC(=O)Nc1ccc([R])cc1", "=SUBSTITUTE($B$2, \"[R]\", A2)", "=SMILES2IMG(C2)", "=SMILES2MW(C2)", "=SMILES2LIPINSKI(C2)"], desc: "4-OH analog (Paracetamol)", out: "151.16 · PASS" },
    { cells: ["F", "", "=SUBSTITUTE($B$2, \"[R]\", A3)", "=SMILES2IMG(C3)", "=SMILES2MW(C3)", "=SMILES2LIPINSKI(C3)"], desc: "4-Fluoro analog", out: "153.16 · PASS" },
    { cells: ["C(F)(F)F", "", "=SUBSTITUTE($B$2, \"[R]\", A4)", "=SMILES2IMG(C4)", "=SMILES2MW(C4)", "=SMILES2LIPINSKI(C4)"], desc: "4-CF3 analog", out: "203.16 · PASS" }
  ]'
/>

* **Next:** [**SMILES, InChI**](/excel/structures) · [**Lipinski, ADME, 3D**](/excel/lipinski-adme)

---

## 6. Dynamic Lead Comparison & FILTER {#dynamic-comparison}

Because `=SMILES2IMG` returns a native **in-cell `CellImage`** on **Microsoft 365 / Excel 2024+**, native lookup and dynamic array formulas move the rendered structures themselves alongside numbers.

<FeatureMedia id="excel-formula-table" />

1. **Dropdown Lead Comparator:** Use Excel **Data Validation** dropdowns in `B2` and `C2` pointing to compound names on a `Library` sheet, then pull each live structure cell with `=XLOOKUP(B2, Library!$A:$A, Library!$C:$C)` paired with `=SMILES.SIMILARITY(B3, C3)` and `ΔMW` / `ΔLogP` rows.
2. **Live Spilling `FILTER` Shortlist:** Combine property cutoffs into a dynamic shortlist whose images update automatically whenever thresholds change:
   ```text
   =FILTER(Library!A2:F200, (Library!E2:E200="PASS")*(Library!F2:F200<=90))
   ```

::: info Microsoft 365 / Excel 2024+ Only
Moving rendered structure images through `XLOOKUP` or `FILTER` requires native in-cell pictures (`=SMILES2IMG`), not Excel 2016–2021 `.FLOAT` shapes.
:::

* **Next:** [**SMILES, InChI**](/excel/structures) · [**Substructure & Similarity (SAR)**](/excel/sar)

---

## Next Steps

- [**Practice in the 10-Tab Excel Playbook (.xlsx)**](/excel/playbook) — Open live workbooks pre-loaded with all formulas above.
- [**10-Second Install Guide**](/guide/install) — Install across Excel, Word, and PowerPoint with one PowerShell command.
