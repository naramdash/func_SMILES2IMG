# Publishing & Export: Batch Style, Vector SVG & Property Tables

Master the **3 essential publishing and export tools** shared across **Excel, Word, and PowerPoint** when preparing journal manuscripts, patent filings, conference posters, and slide decks.

---

## 1. 🎨 Restyle Your Entire Document in 1 Click (`Batch Style`)

When submitting a paper to a print journal or switching a presentation to a dark theme, you never need to redraw dozens of molecules one by one. Click **`[SMILES] ➔ [Batch Style]`** on the top ribbon to convert all chemical structures in your selection (or entire sheet/document/slide deck) in one click!

<BatchStylePreview />

### How Scope Works in Each App
* **📗 Excel:** Converts all chemical formulas in your highlighted cell range—or across the **entire active worksheet** if only a single cell is selected.
* **📘 Word:** Converts structures within your highlighted text/table selection—or across the **entire Word document** if nothing is highlighted.
* **📙 PowerPoint:** Converts your selected slide shapes—or every chemical structure across the **entire presentation** if no shape is selected.

---

## 2. 🖼️ Copy Lossless Vector Graphics for Illustrator & Posters (`Copy SVG`)

Standard screenshots (PNG/JPG bitmaps) blur or pixelate when zoomed in on large conference posters or high-DPI journal proofs.  
Clicking **`[SMILES] ➔ [Copy SVG]`** (or right-clicking ➔ **`Chemical Structure ➔ Copy Vector Graphic (SVG)`**) places a **mathematically sharp scalable vector graphic (`SVG`)** onto your clipboard.

| Target Application | How It Works & Why It's Better |
| :--- | :--- |
| **Adobe Illustrator / Inkscape** | Press <kbd>Ctrl</kbd> + <kbd>V</kbd> to paste editable vector paths and atom labels directly onto your artboard—customize bond thickness, colors, or layout freely. |
| **Conference Posters (A0 Print)** | Scale structures up to any poster size with zero pixelation or jagged edges. |
| **Code & Web Editors (VS Code, HTML)** | Because both the `image/svg+xml` stream and raw `<svg>...</svg>` XML text are copied simultaneously, pasting into a text editor gives you clean inline SVG markup immediately. |

---

## 3. 🔬 Molecule Info Window & 5 Multi-Format Exports

**Right-click** any chemical cell or picture ➔ **`Chemical Structure ➔ Copy Molecule Info`** (or click **`[Molecule Info]`** on the ribbon) to inspect 10+ calculated properties and copy them in the exact format your document needs.

### 10+ Calculated Properties at a Glance
1. **Substance Name:** Common IUPAC / PubChem name (`Aspirin`)
2. **Chemical Formula:** Standard Hill notation (`C9H8O4`)
3. **Molecular Weight (MW):** Standard weight (`180.16 g/mol`)
4. **Monoisotopic Exact Mass:** High-resolution MS mass (`180.0423 Da`)
5. **LogP & TPSA:** Octanol/water partition (`1.31`) & polar surface area (`63.60 Å²`)
6. **Lipinski Rule of 5:** Oral bioavailability status (`PASS`/`FAIL`) plus HBD (`1`), HBA (`4`), and Rotatable Bonds (`3`)
7. **Standard Identifiers:** Canonical SMILES, IUPAC InChI, and 27-character InChIKey

### 📋 5 Copy Formats Tailored to Where You're Pasting

#### ① Auto-Split Across Excel Columns or Word Tables (`Copy as Excel Row — TSV Columns`)
Copies tab-separated values (plus an HTML table stream). **Pressing <kbd>Ctrl</kbd> + <kbd>V</kbd> inside an Excel row or Word/PowerPoint table automatically splits each property into its own adjacent cell!**
```text
Aspirin	C9H8O4	180.16 g/mol	180.0423 Da	CC(=O)Oc1ccccc1C(=O)O	BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ② PowerPoint Slide Bullet Points (`Copy as Multi-Line List`)
Paste straight into any presentation text box for an instant bulleted summary:
```text
• Name: Aspirin
• Formula: C9H8O4
• Molecular Weight: 180.16 g/mol
• Exact Mass: 180.0423 Da
• LogP / TPSA: 1.31 / 63.60 Å²
• Lipinski Ro5: PASS (HBD: 1, HBA: 4, RotBonds: 3)
• SMILES: CC(=O)Oc1ccccc1C(=O)O
• InChIKey: BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ③ Notion · GitHub · Electronic Lab Notebooks (`Copy as Markdown Table`)
```markdown
| Property | Value |
| :--- | :--- |
| **Substance Name** | Aspirin |
| **Chemical Formula** | C9H8O4 |
| **Molecular Weight** | 180.16 g/mol |
| **Exact Mass** | 180.0423 Da |
| **LogP / TPSA** | 1.31 / 63.60 Å² |
| **Lipinski Ro5** | PASS (HBD: 1, HBA: 4, RotBonds: 3) |
```

#### ④ Manuscript Footnotes & Emails (`Copy All Summary`)
```text
Aspirin | Formula: C9H8O4 | MW: 180.16 g/mol | Exact Mass: 180.0423 Da | SMILES: CC(=O)Oc1ccccc1C(=O)O
```

#### ⑤ Python & Data Pipelines (`Export as JSON`)
Inside the **Detailed Molecule Info Window**, choose **JSON** from the export button to copy a structured JSON object containing all calculated descriptors.

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📐 Formula & Option Reference</div>
    <p class="s2i-card-desc">Interactive Excel formula builder and complete cheat sheet of 25+ chemical functions</p>
  </a>
  <a href="../guide/faq" class="s2i-card">
    <div class="s2i-card-title">❓ FAQ & Troubleshooting</div>
    <p class="s2i-card-desc">Quick answers for installation, Excel formulas, styling, SVG copy, and 100% offline security</p>
  </a>
</div>

