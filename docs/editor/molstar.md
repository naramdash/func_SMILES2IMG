---
prev:
  text: 2D Sketcher & Reaction Schemes
  link: /editor/
next:
  text: ACS B&W Batch Style, SVG, Property Tables
  link: /editor/publishing
---

# Insert & Rotate 3D Proteins in 360° (Mol*)

Type a 4-letter PDB code (e.g., `1CRN`, `4HHB`, `6LU7`) or open a local structure file to embed **interactive 360°-rotatable 3D protein models (`.glb`)** or high-resolution 2D figures directly inside PowerPoint, Word, and Excel.

<FeatureMedia id="common-molstar-3d" />

::: tip 💡 Key Highlights of the Built-in 3D Protein Viewer (Mol*)
* **Direct 4-Letter PDB Lookup:** Type `1CRN`, `4HHB`, or `6LU7` to fetch the crystal structure immediately, or open local `.pdb`, `.cif`, and `.sdf` files.
* **Live 360° Rotation on Slides (`.glb`):** Drag with your mouse right on your PowerPoint slide, Word page, or Excel sheet to rotate and zoom the protein in 3D.
* **Double-Click Re-Editing Across PCs:** **Double-click** any inserted 3D model or 2D picture to reopen the viewer at the exact same camera angle and style!
:::

## 1. Insert a 3D Protein in 3 Steps

### Step 1. Open the Mol* Viewer
* Click **`[3D Structure (Mol*)]`** on the **`[SMILES]`** ribbon tab (or right-click ➔ **`Chemical Structure ➔ 3D Structure (Mol*)...`**).
* 💡 **Excel Shortcut:** If your selected Excel cell already contains a 4-letter PDB code like `1CRN` or `4HHB`, opening Mol\* loads that structure automatically!

### Step 2. Load & Style Your Structure
1. Enter a **4-character PDB ID** in the top bar and click **`Fetch`**, or click **`Open File...`** to load a local `.pdb`, `.cif`, `.bcif`, `.sdf`, `.mol`, or `.mol2` file.
2. Customize the visual style:
   * **Cartoon / Ribbon:** Highlight α-helices, β-sheets, and loops.
   * **Ball & Stick / Surface:** Inspect binding pockets, active-site residues, and cofactors.
   * **Color Schemes:** Color by polymer chain, secondary structure, residue type, or B-factor.

### Step 3. Insert into Your Document (3D Model vs. 2D Picture)
Select the output mode in the top toolbar and click **`Apply to Slide`** / **`Apply to Document`** (or press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>):

| Insert Mode | Best For |
| :--- | :--- |
| **🧊 `3D Model (GLB)`** | Embeds an **interactive 3D object** you can drag with your mouse to rotate 360° during PowerPoint slide shows or inside Excel and Word. |
| **🖼️ `2D Picture`** | Exports a **high-resolution 2D `.png` figure** for print journals and PDFs (use **`BG: Transparent`** for dark slide themes, or toggle White / Dark backgrounds). |

## 2. Sample PDB Codes to Try Right Now

Not sure which code to test first? Try any of these 4 classic structures:

| PDB ID | Biomolecule | What to Look For |
| :---: | :--- | :--- |
| **`1CRN`** | **Crambin** | Compact plant protein — fast test for α-helices and β-sheets |
| **`4HHB`** | **Human Hemoglobin** | Tetrameric (α₂β₂) protein complex with 4 bound heme groups |
| **`6LU7`** | **SARS-CoV-2 Main Protease (Mpro)** | Drug discovery example showing a covalent inhibitor bound inside the active site |
| **`1BNA`** | **B-DNA Double Helix** | Classic 12-base-pair DNA duplex showing base pairing and major/minor grooves |

## 3. Double-Click to Adjust Camera or Style Later

<FeatureMedia id="powerpoint-slide-molstar-3d" />

1. Select any previously inserted Mol\* 3D model or 2D picture in Excel, Word, or PowerPoint.
2. **Double-click** it (or right-click ➔ **`Chemical Structure ➔ 3D Structure (Mol*)...`**).
3. Mol\* reopens with your exact protein, custom colors, representations, and camera angle intact.
4. Tweak the view and click **Apply** to update the figure in place without changing its size or position!

## 🚀 Continue Exploring

- [**PowerPoint Presentations**](/powerpoint/) — Rotate 3D protein models live during slide shows and style dark-mode structures
- [**Ketcher 2D Sketcher**](/editor/) — Draw small-molecule ligands and multi-step reaction schemes directly inside Office
