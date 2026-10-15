# Mol* 3D Macromolecular Viewer (Excel, Word & PowerPoint)

**SMILES for Microsoft Office Suite** includes a fully embedded, 100% offline **Mol\* 3D Macromolecular Viewer** (v5.13.1) powered by Microsoft Edge WebView2.

---

## 🌟 Overview & Capabilities

Mol\* (pronounced *"mol-star"*) is the modern open-source standard for visualizing biological macromolecules, protein structures, nucleic acids, and small-molecule complexes.

* **100% Local & Offline:** The entire viewer bundle is packaged directly inside the add-in. Structures render with hardware-accelerated WebGL entirely on your local GPU without cloud servers.
* **Dual Output Modes:**
  * **Interactive 3D Model (`.glb`):** Insert genuine, rotatable, zoomable 3D models directly into PowerPoint slides, Word documents, and Excel worksheets!
  * **Publication 2D Image (`.png`):** Export crisp, high-resolution 2D figures with optional transparent background for print publications, presentations, and spreadsheets.
* **Double-Click In-Place Re-editing:** Double-click any inserted 3D model or 2D picture in Excel, PowerPoint, or Word to immediately re-open Mol\* in the exact camera angle, zoom level, and representation style!
* **Lossless Session Persistence:** Complete viewer session recipes (`.molj`) are stored losslessly inside the document's `CustomXMLParts`. When sharing workbooks, presentations, or documents with colleagues on other computers, double-clicking restores the exact 3D workspace.

---

## 🚀 Getting Started with Mol*

### 1. Launching Mol*
* **Ribbon Tab:** Go to the **`[SMILES]`** tab in Excel, PowerPoint, or Word and click **`[3D Structure (Mol*)]`**.
* **Context Menu:** Right-click any slide, document area, or Excel cell ➔ **`3D Structure (Mol*)...`**.
* **Smart PDB Preloading (Excel):** If your active Excel cell contains a 4-character PDB ID (e.g. `1CRN`, `4HHB`) or structure file path, Mol\* automatically preloads that structure!

### 2. Loading Structures
Inside the Mol\* viewer window:
* **Direct PDB ID Download:** Enter any 4-character RCSB PDB ID (e.g. `1CRN` Crambin, `4HHB` Hemoglobin, `7BWY` SARS-CoV-2 Spike) and click **Download PDB**.
* **Local Structure Files:** Click **Open File** to load structures in `.pdb`, `.cif`, `.bcif`, `.sdf`, `.mol`, or `.mol2` format.
* **Saved Sessions:** Open previously saved `.molj` / `.molx` Mol\* session recipes.

### 3. Customizing Visual Representations
Use the built-in representation presets:
* **Cartoon / Ribbon:** Emphasizes secondary structures ($\alpha$-helices, $\beta$-sheets, loops).
* **Ball & Stick / Molecular Surface:** Details binding pockets, active site residues, and cofactor ligands.
* **Color Schemes:** Color by Polymer Chain, Secondary Structure, Residue Type, or B-factor / Uncertainty.

---

## 📦 Inserting into Documents: 3D Model vs. 2D Picture

In the top action bar of the Mol\* window, choose your desired target format:

### Mode A: [Insert 3D Model (.glb)] (Interactive 3D)
1. Mol\* exports the current camera view and representations as a binary GLTF (`.glb`) asset.
2. The model is embedded directly onto your active slide, document, or Excel worksheet (aligned to the active cell with `xlMoveAndSize`).
3. **Interactive Control in PowerPoint & Excel:**
   - Click and drag the 3D control handle in the center of the model to freely rotate, pitch, and zoom the macromolecule in 360 degrees directly inside your slide or spreadsheet!

### Mode B: [Insert 2D Picture (.png)] (Publication Quality)
1. Renders the current viewport as a high-resolution PNG image.
2. Choose **Transparent Background** or **Solid Background** to blend seamlessly onto dark slide templates, white document pages, or spreadsheet cells.

---

## 🔄 Seamless Re-editing Workflow

1. Click any previously inserted Mol\* 3D model or 2D picture in Excel, Word, or PowerPoint.
2. **Double-click** the object (or right-click ➔ **`Edit 3D Structure (Mol*)...`**).
3. The Mol\* editor instantly opens, restoring the exact macromolecule, custom colors, representations, and camera view.
4. Adjust the perspective or representation and click **Apply** to update the object in-place without altering position or dimensions.
