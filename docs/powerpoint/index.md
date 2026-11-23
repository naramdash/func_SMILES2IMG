# PowerPoint Overview

Create transparent dark-mode structures, reaction pathways, rotatable 3D proteins (`Mol*`), and scalable vector SVGs for A0 posters directly inside Microsoft PowerPoint.

<FeatureMedia id="powerpoint-overview" />

## Key Workflows

### 1. Transparent Dark Mode & Deck-Wide Theme Switching (`Batch Style`)
Move dark conference slides into a bright lab-meeting template or prepare monochrome handouts without redrawing shapes. Switch **2D structures across the presentation** between Dark Mode, Standard Color, ACS B&W, and White Card using `[Batch Style]`—keeping coordinates and alignment intact.

<FeatureMedia id="powerpoint-slide-batch-style" />

### 2. Interactive 360° Protein Rotation & Morph Zoom-In (`Mol*`)
Insert a 4-character PDB ID (`1CRN`, `6LU7`) or local `.pdb`/`.cif` file onto a slide as a native **3D `.glb` model**—rotate binding pockets with the mouse during talks, or animate smooth camera zoom-ins using PowerPoint **Morph** without external video exports.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

### 3. 4-Card Candidate Grids & TextBox Conversion
Select a TextBox containing a SMILES string or chemical name (`Aspirin`, `Paracetamol`), click **`[Paste Structure]`**, switch to transparent white ink, and paste formatted `Molecule Info` bullets underneath.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

## Comparison: Chemistry Presentations & Posters

| Presentation Task | Without `func_SMILES2IMG` | With `func_SMILES2IMG` |
| :--- | :--- | :--- |
| **Placing structures on dark navy/slate slides** | Opaque white boxes or manual background-removal steps | **Alpha transparency + white ink (`Batch Style: Dark Mode`)** |
| **Changing slide template theme** | Re-export and realign PNGs slide by slide | **Clear selection and run `[Batch Style]` across the deck** |
| **Showing a 3D protein binding pocket** | Multiple static PyMOL angles or pre-recorded MP4s | **Native `3D Model (GLB)` rotatable in Slide Show + Morph zoom-in** |
| **Presenting on a podium PC or macOS** | Proprietary OLE objects or missing fonts | **Standard PNG + `.glb` embedded—works without installing the add-in** |
| **Printing large A0 conference posters** | Bitmaps pixelate when scaled up | **`[Copy SVG]` copies scalable vector paths** |

## Podium & Lab Compatibility

<OfficeOverviewCta app="powerpoint" />

## Quick Walkthrough

1. **Convert a TextBox:** Insert a TextBox on any slide, type `Aspirin` or `CC(=O)Oc1ccccc1C(=O)O`, select the TextBox, and click **`[SMILES]` ➔ `[Paste Structure]`** (or right-click ➔ **`Chemical Structure` ➔ `Paste Structure`**).
2. **Switch to Dark Mode & re-edit in place:** On a dark slide, click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode]`**, then **double-click** the structure to open **Ketcher 2D** and edit in place.
3. **Practice inside the 8-slide Playbook (`.pptx`):** Download <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_PowerPoint_Playbook.pptx" download="func_SMILES2IMG_PowerPoint_Playbook.pptx"><strong>func_SMILES2IMG_PowerPoint_Playbook.pptx</strong></a> and open it with Speaker Notes visible to practice TextBox conversion, 4-card grids, KET reactions, Mol\* `1CRN`, and deck-wide restyling ([see 8-slide breakdown →](/powerpoint/playbook)).

## Jump to a PowerPoint Workflow

| What You Want to Do in PowerPoint | Guide | Key Command |
| :--- | :--- | :--- |
| See 3D motion, SAR card, A0 poster & animated mechanism scenarios | [PowerPoint Use Cases](/powerpoint/use-cases) | 5 stage & poster scenarios |
| Convert TextBoxes in-place, draw 2D molecules, and build `KET` reaction slides | [1. 2D Structures & Reaction Slides](/powerpoint/slides-structures) | `[Paste Structure]` / `[Insert / Edit Structure]` |
| Style transparent white-ink cards, build 4-card SAR grids, and switch deck themes | [2. Dark Mode, Cards, Deck Restyle](/powerpoint/dark-mode) | `[Batch Style]` + `[Molecule Info]` |
| Embed 360°-rotatable 3D protein models (`.glb`) and animate binding pockets with **Morph** | [3. 360° 3D Protein Slides (Mol* & Morph)](/powerpoint/3d-proteins) | `[3D Structure (Mol*)]` |
| Copy vector `SVG` for A0 posters, prep Q&A property notes, and check podium compatibility | [4. A0 Posters (SVG), Q&A, Podium Sharing](/powerpoint/publishing-qa) | `[Copy SVG]` / `[Molecule Info]` |
| Inspect every `[SMILES]` ribbon button and the 5 slide right-click context menus | [Ribbon & Context Menus](/powerpoint/reference) | Top `[SMILES]` tab & Right-Click |
