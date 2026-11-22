---
prev:
  text: 3D Protein Viewer & GLB Models
  link: /editor/molstar
next:
  text: WebView2 Loading
  link: /editor/troubleshooting/webview2-loading
---

# 1-Click ACS Journal B&W, Vector SVG, Property Table Export

Convert every chemical structure in an Excel sheet, Word manuscript, or PowerPoint deck into **ACS Black & White** or **Dark Mode** in a single click—and copy lossless **SVG vectors** straight into Adobe Illustrator.

## 1. 🎨 Restyle Your Entire Document in 1 Click (`Batch Style`)

When submitting a paper to a print journal or switching a presentation to a dark theme, you never need to redraw dozens of molecules one by one. Click **`[SMILES]`** ➔ **`[Batch Style]`** on the top ribbon to convert all chemical structures in your selection (or entire sheet/document/slide deck) in one click:

| `Batch Style` Preset | Background & Bond Style | Best For | Equivalent Excel Formula |
| :--- | :--- | :--- | :--- |
| **🎨 `Standard Color (Transparent)`** | Transparent BG + CPK heteroatom colors | Everyday reports & light slides | `=SMILES2IMG(A2)` |
| **✒️ `Publication B&W (ACS Monochrome)`** | Transparent BG + pure black line art | Journal submissions (ACS style) & print | `=SMILES2IMG(A2, "", "bw")` |
| **⬜ `White Background (Color)`** | Opaque white rectangle + CPK colors | Shaded table rows or colored backgrounds | `=SMILES2IMG(A2, "white")` |
| **⬛ `Dark Mode (White Ink)`** | Transparent BG + crisp white ink | Dark-themed conference slide decks | `=SMILES2IMG(A2, "trans", "white")` |

### How Scope Works in Each App
* **📗 Excel:** Converts all chemical formulas in your highlighted cell range—or across the **entire active worksheet** if only a single cell is selected.
* **📘 Word:** Converts structures within your highlighted text/table selection—or across the **entire Word document** if nothing is highlighted.
* **📙 PowerPoint:** Converts your selected slide shapes—or every chemical structure across the **entire presentation** if no shape is selected.

<FeatureMedia id="word-ch6-publication-finish" />

## 2. 🖼️ Copy Lossless Vector Graphics for Illustrator & Posters (`Copy SVG`)

Standard screenshots (PNG/JPG bitmaps) blur or pixelate when zoomed in on large conference posters or high-DPI journal proofs.  
Clicking **`[SMILES] ➔ [Copy SVG]`** (or right-clicking ➔ **`Chemical Structure ➔ Copy Vector SVG`**) places a **mathematically sharp scalable vector graphic (`SVG`)** onto your clipboard.

| Target Application | How It Works & Why It's Better |
| :--- | :--- |
| **Adobe Illustrator / Inkscape** | Press <kbd>Ctrl</kbd> + <kbd>V</kbd> to paste editable vector paths and atom labels directly onto your artboard—customize bond thickness, colors, or layout freely. |
| **Conference Posters (A0 Print)** | Scale structures up to any poster size with zero pixelation or jagged edges. |
| **Code & Web Editors (VS Code, HTML)** | Because both the `image/svg+xml` stream and raw `<svg>...</svg>` XML text are copied simultaneously, pasting into a text editor gives you clean inline SVG markup immediately. |

## 3. 🔬 Molecule Info Window & 5 Multi-Format Exports

**Right-click** any chemical cell or picture ➔ **`Chemical Structure ➔ Copy Molecule Info`** (or click **`[Molecule Info]`** on the ribbon) to inspect calculated properties and copy them in the exact format your document needs. Clicking **`Detailed Molecule Info Window...`** opens the standard Windows dialog with a live 2D structure thumbnail preview, 1-click row **`Copy`** buttons, and a bottom format selector:

<FeatureMedia id="common-molecule-info" />

### 10+ Calculated Properties at a Glance
1. **Substance Name:** Common IUPAC / PubChem name (`Aspirin`)
2. **Chemical Formula:** Standard Hill notation (`C9H8O4`)
3. **Molecular Weight (MW):** Standard weight (`180.16 g/mol`)
4. **Monoisotopic Exact Mass:** High-resolution MS mass (`180.0423 Da`)
5. **LogP & TPSA:** Octanol/water partition (`1.31`) & polar surface area (`63.60 Å²`)
6. **Lipinski Rule of 5:** Oral bioavailability status (`PASS`/`FAIL`) plus HBD (`1`), HBA (`4`), and Rotatable Bonds (`3`)
7. **Rings, Aromaticity & 3D Complexity:** Heavy Atoms (`13`), Total Rings (`1`), Aromatic Rings (`1`), Fraction Csp3 (`Fsp3: 0.11`), and Chiral Centers (`0`)
8. **Standard Identifiers:** Canonical SMILES, IUPAC InChI, and 27-character InChIKey

### 📋 5 Copy Formats Tailored to Where You're Pasting

#### ① Start a New Table — 17-Column Header + Data Row (`Copy All Summary (Text)`)
Copies a **2-row table (17 headers on Row 1 + 17 values on Row 2)** as both TSV text and a formatted HTML table. Paste into an empty Excel range or Word page to create a complete 17-column table in one step:
```text
Name	Formula	MW (g/mol)	Exact Mass	LogP	TPSA (Å²)	HBD	HBA	RotB	Heavy Atoms	Rings	Aromatic Rings	Fsp3	Chiral Centers	Lipinski	SMILES	InChIKey
Aspirin	C9H8O4	180.16	180.0423	1.31	63.6	1	4	3	13	1	1	0.11	0	Pass	CC(=O)Oc1ccccc1C(=O)O	BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ② Append to an Existing Table — 17-Column Data Row Only (`Copy as Excel Row (TSV Columns)`)
Copies **only the 17 tab-separated values without headers**. **Pressing <kbd>Ctrl</kbd> + <kbd>V</kbd> in the next empty row of an Excel sheet or Word/PowerPoint table automatically splits each property into its own adjacent cell!**
```text
Aspirin	C9H8O4	180.16	180.0423	1.31	63.6	1	4	3	13	1	1	0.11	0	Pass	CC(=O)Oc1ccccc1C(=O)O	BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ③ PowerPoint Slides & Experimental Summaries (`Copy as Multi-Line List`)
Paste straight into any presentation text box or report for a clean multi-line summary:
```text
Substance Name: Aspirin
IUPAC Name: 2-acetyloxybenzoic acid
Formula: C9H8O4
Molecular Weight: 180.16 g/mol
Exact Mass: 180.0423
LogP: 1.31
TPSA: 63.6 Å²
H-Bond Donors: 1
H-Bond Acceptors: 4
Rotatable Bonds: 3
Heavy Atoms: 13
Rings (Total / Aromatic): 1 / 1
Fraction Csp3 (Fsp3): 0.11
Chiral Centers: 0
Lipinski Rule of 5: Passed
SMILES: CC(=O)Oc1ccccc1C(=O)O
InChI: InChI=1S/C9H8O4/c1-6(10)13-8-5-3-2-4-7(8)9(11)12/h2-5H,1H3,(H,11,12)
InChIKey: BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ④ Notion · GitHub · Electronic Lab Notebooks (`Copy as Markdown Table`)
```markdown
| Property | Value |
| :--- | :--- |
| **Substance Name** | Aspirin |
| **IUPAC Name** | 2-acetyloxybenzoic acid |
| **Formula** | C9H8O4 |
| **Molecular Weight** | 180.16 g/mol |
| **Exact Mass** | 180.0423 |
| **LogP** | 1.31 |
| **TPSA** | 63.6 Å² |
| **HBD / HBA / RotB** | 1 / 4 / 3 |
| **Heavy Atoms / Chiral** | 13 / 0 |
| **Rings (Total / Aromatic)** | 1 / 1 |
| **Fraction Csp3 (Fsp3)** | 0.11 |
| **Lipinski Rule of 5** | Passed |
| **SMILES** | `CC(=O)Oc1ccccc1C(=O)O` |
| **InChI** | `InChI=1S/C9H8O4/c1-6(10)13-8-5-3-2-4-7(8)9(11)12/h2-5H,1H3,(H,11,12)` |
| **InChIKey** | `BSYNRYMUTXBXSQ-UHFFFAOYSA-N` |
```

#### ⑤ Python & Data Pipelines (`Export as JSON`)
Inside the **Detailed Molecule Info Window**, select **JSON** from the format dropdown at the bottom and click **`Copy All`** (or right-click any row ➔ **`Copy All (JSON)`**) to copy a structured JSON object containing all calculated descriptors:
```json
{
  "name": "Aspirin",
  "iupacName": "2-acetyloxybenzoic acid",
  "formula": "C9H8O4",
  "molecularWeight": 180.16,
  "exactMass": 180.0423,
  "logP": 1.31,
  "tpsa": 63.6,
  "hBondDonors": 1,
  "hBondAcceptors": 4,
  "rotatableBonds": 3,
  "heavyAtomCount": 13,
  "ringCount": 1,
  "aromaticRingCount": 1,
  "fsp3": 0.11,
  "chiralCenters": 0,
  "passesLipinski": true,
  "smiles": "CC(=O)Oc1ccccc1C(=O)O",
  "inchi": "InChI=1S/C9H8O4/c1-6(10)13-8-5-3-2-4-7(8)9(11)12/h2-5H,1H3,(H,11,12)",
  "inchiKey": "BSYNRYMUTXBXSQ-UHFFFAOYSA-N"
}
```

## 🚀 Continue Exploring

- [**All 40 Excel Functions**](/excel/formulas) — Complete parameter catalog, data types, and shared rendering options for all 40 chemical functions
- [**FAQ & Troubleshooting**](/guide/faq) — Quick answers for installation, Excel formulas, styling, SVG copy, and local-first privacy
