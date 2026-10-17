# Excel Formula Generator & Style Simulator

Build and preview custom Excel chemical formulas interactively without memorizing option keywords. Click any combination of options below to **preview the rendered 2D structure in real time and copy the ready-to-use formula in one click**.

---

## 1. 🧪 Interactive Excel Formula Expression Generator

Select your input data type (`SMILES`, `Chemical Name`, `CAS Number`, `Reaction KET`, `InChI`), Excel version, background color, bond ink mode, substructure highlight (`hl:`), and atom labels (`num`):

<FormulaBuilder />

---

## 2. 🎨 1-Click `Batch Style` Preset Simulator

Preview the 4 whole-document style presets available under **`[SMILES]`** ➔ **`[Batch Style]`** on the Excel, Word, and PowerPoint ribbon:

<BatchStylePreview />

---

## 3. ⚡ Quick-Copy Property, Lipinski & SAR Formulas

Pair your structure column with these zero-latency local calculation formulas (assuming `A2` holds your molecule):

| Property to Calculate | Excel Formula to Copy | Example Output (Aspirin) |
| :--- | :--- | :--- |
| **Chemical Formula (Hill)** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **Molecular Weight (MW, g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **Monoisotopic Exact Mass (Da)** | `=SMILES2EXACTMASS(A2)` | `180.0423` |
| **Lipinski Rule of 5 Pass/Fail** | `=SMILES2LIPINSKI(A2)` | `PASS` |
| **LogP / TPSA (Å²)** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | `1.31` / `63.60` |
| **Substructure Match (e.g., Carboxylic Acid)** | `=SMILES.MATCH(A2, "[CX3](=O)[OX2H1]")` | `TRUE` |
| **Tanimoto Similarity vs. Reference (`$D$1`)** | `=SMILES.SIMILARITY(A2, $D$1)` | `0.85` (0.0 – 1.0) |

---

## 🔗 Related Guides

<div class="feature-grid-2">
  <div class="feature-card">
    <h4>📐 <a href="/reference/formula">Excel Formula Cheat Sheet</a></h4>
    <p>Full parameter specifications and alias tables for all 25+ Excel chemical functions.</p>
  </div>
  <div class="feature-card">
    <h4>📗 <a href="/guide/excel">Microsoft Excel Basics Guide</a></h4>
    <p>Learn how to build compound tables, resize structure cells, and edit molecules in-place.</p>
  </div>
</div>
