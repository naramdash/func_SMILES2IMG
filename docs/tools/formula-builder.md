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
| **Heavy Atoms / Rings / Aromatic Rings** | `=SMILES2HEAVY(A2)` / `=SMILES2RINGS(A2)` / `=SMILES2AROMATIC(A2)` | `13` / `1` / `1` |
| **3D Complexity (Fsp³ / Chiral Centers)** | `=SMILES2FSP3(A2)` / `=SMILES2CHIRAL(A2)` | `0.11` / `0` |
| **Substructure Match (e.g., Carboxylic Acid)** | `=SMILES.MATCH(A2, "[CX3](=O)[OX2H1]")` | `TRUE` |
| **Tanimoto Similarity vs. Reference (`$D$1`)** | `=SMILES.SIMILARITY(A2, $D$1)` | `0.85` (0.0 – 1.0) |

---

## 4. 🔍 Quick-Copy Dictionary: 12 Common Functional Group SMARTS (`MATCH` & Highlight)

Copy any formula below to filter rows containing a specific functional group (`=SMILES.MATCH`) or highlight that group directly inside the rendered structure (`=SMILES2IMG`):

| Functional Group / Ring | Boolean Filter (`TRUE` / `FALSE`) | Highlighted Structure Image Formula |
| :--- | :--- | :--- |
| **Benzene Ring (Aromatic)** | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, , "hl:c1ccccc1")` |
| **Pyridine Ring** | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, , "hl:c1ccncc1")` |
| **Carboxylic Acid (`-COOH`)** | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, , "hl:C(=O)[OH]")` |
| **Ester (`-COO-`)** | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, , "hl:C(=O)O[#6]")` |
| **Amide (`-CONH-`)** | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, , "hl:C(=O)N")` |
| **Sulfonamide (`-SO₂NH-`)** | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, , "hl:S(=O)(=O)N")` |
| **Ketone (`-C(=O)-`)** | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, , "hl:[#6]C(=O)[#6]")` |
| **Aldehyde (`-CHO`)** | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, , "hl:[CX3H1](=O)")` |
| **Phenol (Aromatic `-OH`)** | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, , "hl:c1ccccc1[OH]")` |
| **Nitro Group (`-NO₂`)** | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, , "hl:[N+](=O)[O-]")` |
| **Any Halogen (`F, Cl, Br, I`)** | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, , "hl:[F,Cl,Br,I]")` |
| **Trifluoromethyl (`-CF₃`)** | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, , "hl:C(F)(F)F")` |

---

## 5. 📊 8-Tab Master Interactive Excel Showcase (`SMILES2IMG_Comprehensive_Showcase.xlsx`)

Want to explore all formulas above inside a live Excel workbook without typing? Download the official **8-tab master showcase workbook**:

<div class="s2i-card-grid" style="margin: 14px 0 18px 0;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 Download 8-Tab Master Showcase Workbook (.xlsx)</div>
    <p class="s2i-card-desc">Pre-built with live in-cell structures, Lipinski Ro5 drug screening, substructure matching, styling matrix, and KET reaction canvases (0 formula errors verified)</p>
  </a>
</div>

| Sheet Tab Name | What You Can Practice & Copy | Key Formulas Demonstrated |
| :--- | :--- | :--- |
| **`0. Quick Guide`** | Keyboard shortcuts (`Ctrl+Shift+K`), double-click editing workflow, and navigation index | Overview & workflow tips |
| **`1. Structure to Image`** | Rendering 2D structures from SMILES, chemical names, CAS numbers, and InChI strings | `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=NAME2SMILES` |
| **`2. Molecular Properties`** | Drug discovery SAR & 3D complexity screening (Lipinski Ro5, Rings, Fsp3, Chiral) | `=SMILES2FORMULA`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2RINGS`, `=SMILES2FSP3`, `=SMILES2CHIRAL`, `=SMILES2LIPINSKI` |
| **`3. Substructure & Similarity`** | Functional group filtering (`SMARTS`), scaffold highlighting, and Tanimoto similarity | `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2IMG(..., "hl:...")` |
| **`4. Drawing Styles & Options`** | Side-by-side matrix of ACS monochrome, dark mode, custom ink colors, atom numbering, and flips | `"bw"`, `"white"`, `"ink=#1e40af"`, `"num stereo"`, `"flip rot90"` |
| **`5. Reactions & KET Canvas`** | Multi-step synthesis schemes and lossless KET JSON structure rendering | `=KET2IMG`, double-click reaction canvas editing |
| **`6. Legacy Excel (.FLOAT)`** | Floating picture formulas designed for Excel 2016, 2019, and 2021 compatibility | `=SMILES2IMG.FLOAT`, `=NAME2IMG.FLOAT`, `=KET2IMG.FLOAT` |
| **`7. 3D & Ribbon Tools`** | Mol\* 3D protein viewer (`1CRN`, `4HHB`), `Copy SVG`, `Batch Style`, and `.sdf` import/export | `[SMILES]` Ribbon & Right-Click Context Menu guide |

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
