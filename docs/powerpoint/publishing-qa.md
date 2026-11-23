# 4. A0 Posters (SVG), Q&A, Podium Sharing

Run this pre-conference checklist to prepare scalable **A0 / A1 poster graphics**, **live Q&A chemistry notes**, and reliable playback on **conference podium PCs or Macs**.

<FeatureMedia id="powerpoint-slide-live-qa-info" />

## Conference & Poster Checklist

### Step 1: Copy Vector `SVG` for A0 / A1 Posters (`Copy SVG`)

Standard PNG or JPG captures look fine on a laptop screen, then turn jagged and blurry when printed at **A0 / A1 poster dimensions (90 × 120 cm)** at 300 DPI.

1. Click any 2D molecule or `KET` reaction pathway on a slide (or highlight a SMILES/name text string inside a TextBox).
2. Click **`[SMILES]` ➔ `[Copy SVG]`** (or right-click ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**).
3. Paste (<kbd>Ctrl</kbd> + <kbd>V</kbd>) onto your PowerPoint A0 poster canvas (Office 365 / 2019+) or into **Adobe Illustrator / Inkscape**—every bond path, wedge, and atom letter pastes as editable vector paths ready to scale cleanly.

### Step 2: Prepare Q&A Chemistry Metrics (`Molecule Info`)

In full-screen Slide Show mode, PowerPoint hides the Ribbon and right-click add-in menus. Use this two-step Q&A workflow instead:

* **Before Stepping on Stage:** For key candidates, select each structure shape ➔ click **`[SMILES]` ➔ `[Molecule Info]` ➔ `[Copy as Multi-Line List]`**, and paste Formula, Exact Mass, LogP, TPSA, and `Fsp3` into **Speaker Notes** (visible to you in Presenter View).
* **During Deep-Dive Q&A / Group Meetings:** Press <kbd>Esc</kbd> to return to Normal view, click the structure (or highlight any SMILES/name text), and open **`[Molecule Info]` ➔ `[Detailed Molecule Info Window...]`** to display all 12 descriptors with a 2D preview—double-click any row to copy a single value. *(See [Style & Export Reference](/editor/publishing) for all 5 copy formats.)*

### Step 3: Verify Podium PC, macOS & PDF Handouts

* **Podium PC / macOS Playback:** Both 2D PNGs and 3D `.glb` models (including **Morph** camera transitions) are native PowerPoint shapes embedded inside `.pptx`, so they run cleanly on conference podium PCs or **macOS** machines without `func_SMILES2IMG` installed.
* **Lightweight PDF / Legacy PowerPoint Handouts:**
  * Converting `.pptx` to PDF flattens 3D `.glb` shapes at their resting camera angle—make sure each 3D slide rests at a clear view before exporting.
  * For printed monochrome handouts, save a duplicate copy and run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`**.

## Related Guides

- [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) — Restyle your entire deck theme via Batch Style
- [**3. 360° 3D Protein Slides (`Mol*` & Morph)**](/powerpoint/3d-proteins) — Configure 3D `.glb` models and Morph transitions
- [**Ribbon & Context Menus**](/powerpoint/reference) · [**Troubleshooting**](/powerpoint/troubleshooting/dark-slides) — Command map and quick fixes for dark slides, groups, and `.glb` models
