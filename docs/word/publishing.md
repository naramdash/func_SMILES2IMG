# 4. Journal Finish, Vector SVG, Co-Author Sharing

Run this pre-submission checklist before exporting your Microsoft Word manuscript to PDF or uploading to ACS, RSC, Wiley, Elsevier, Nature, or patent portals.

<FeatureMedia id="word-ch6-publication-finish" />

## Pre-Submission Checklist

### Step 1: Unify Every 2D Structure & Scheme to ACS B&W (`Batch Style`)

When colored internal drafts need strict monochrome line art for journal submission:

1. **Choose Your Scope in Word:**
   * **Entire Manuscript (Inline + Floating):** Leave no picture selected—simply keep your text cursor anywhere in prose.
   * **Single Table or Figure Only:** Highlight the specific table rows or paragraphs you want to convert.
2. Click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`**.
   * `func_SMILES2IMG` re-renders every 2D molecule and `KET` reaction scheme with transparent backgrounds and crisp black line art while preserving exact image sizes and column widths (`Mol*` 3D protein figures are automatically skipped).
   * Need to revert to color for a thesis or color preprint later? Run **`[Batch Style]` ➔ `[Standard Color (Transparent)]`** (see [Style & Export Reference](/editor/publishing) for all 4 preset specs).

### Step 2: Export Vector `SVG` for TOC Graphics & Multi-Panel Figures

Standard bitmaps pixelate when resized or edited inside **Adobe Illustrator**, **Inkscape**, or **CorelDRAW** for a **Table of Contents (TOC) / Graphical Abstract** or multi-panel `Figure 1A–C`:

1. Click any 2D structure or KET reaction scheme in your Word manuscript (or highlight a SMILES string in prose).
2. Click **`[SMILES]` ➔ `[Copy SVG]`** (or right-click ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**).
3. Press <kbd>Ctrl</kbd> + <kbd>V</kbd> inside Illustrator, Inkscape, or modern Microsoft 365 documents—every bond path, wedge, and atom label pastes as editable vector paths.

### Step 3: Extract Supporting Information (SI) Identifiers Without Typos

* **Characterization & HRMS Blocks:** Select each new compound in your manuscript and click **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** to paste clean Hill Formula, Average MW, Monoisotopic Exact Mass (`Da`), LogP, TPSA, SMILES, InChI, and InChIKey lines straight into your Supporting Information.
* **Single Identifiers for Submission Forms:** Use **`[Copy SMILES]`** or **`[Copy InChI]`** on the ribbon whenever journal portals ask for machine-readable compound strings.

### Step 4: Verify Co-Author, macOS & PDF Sharing

* **No Broken Images for Readers:** All 2D structures are embedded as standard high-DPI PNG images and 3D models as standard `.glb` objects. Co-authors on **macOS**, Windows PCs without `func_SMILES2IMG`, Word Online, and journal editorial PDF converters open and print the `.docx` identically.
* **Round-Trip Editing Survives Co-Author Reviews:** Because chemical tags remain stored in picture **Alt Text** and `CustomXMLParts`, your co-authors can edit surrounding text or track changes on Mac, and when you reopen the `.docx` on Windows, double-click editing still works.

::: tip Reuse Your Manuscript Schemes in PowerPoint
Preparing a group meeting or conference deck from your paper? Click any KET reaction scheme in Word, press **`[SMILES]` ➔ `[Copy KET]`**, switch to PowerPoint, click **`[Paste Structure]`**, and run **`[Batch Style]` ➔ `[Dark Mode]`**—your entire synthesis scheme transforms into transparent white ink on your dark slides.
:::

## Related Guides

- [**2. SAR Tables & Experimental Properties**](/word/tables-properties) — Lock table column widths and populate 12 descriptors automatically
- [**Style & Export Reference**](/editor/publishing) — Complete breakdown of all 4 `Batch Style` presets and 5 `Molecule Info` clipboard formats
- [**Document Sharing Guide**](/guide/troubleshooting/document-sharing) — How `.docx`, `.pptx`, and `.xlsx` files behave across macOS and add-in-free PCs
