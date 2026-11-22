# 3D Models, Double-Click & Podium Sharing

Troubleshooting for Mol\* `.glb` 3D protein insertion, `.pptx` file size limits, double-clicking structures or reactions on slides, and offline playback on conference podium PCs and macOS.

## 1. 3D `.glb` Protein Won't Insert, Rotate, or Makes `.pptx` Too Large {#3d-models}

* **Check Office Version Requirement:** Live mouse rotation and **Morph** camera transitions for `.glb` objects require **PowerPoint 2019, 2021, 2024, or Microsoft 365**.
* **Lightweight Decks, PowerPoint 2016, or Giant Cryo-EM/Ribosome Complexes:** Switch Mol\*'s top toolbar from **`3D Model (GLB)`** to **`2D Picture`** with **`BG: Transparent`** to insert a sharp alpha-transparent PNG that works on every PowerPoint version while keeping file size small.
* **Corporate Firewall Blocking RCSB PDB Fetches:** Click **`Open File...`** inside Mol\* to load local `.pdb`, `.cif`, `.bcif`, or `.sdf` files offline ([Offline PDB Files](/editor/troubleshooting/offline-pdb)).

## 2. Double-Click Opens Format Picture Instead of Ketcher or Mol\* {#double-click}

`func_SMILES2IMG` identifies editable slide objects via **`Shape.Tags`** (`KET`, `SMILES`, `MOLSTAR`) and **`Alt Text`**, plus `CustomXMLParts` for 3D Mol\* camera states.

* **Plain Screenshot or Copied Bitmap:** External images lack chemical tags and open PowerPoint Format Picture instead. Re-insert via **`[Insert / Edit Structure]`** or **`[Paste Structure]`**.
* **SVG Converted to Shapes:** Once an SVG pasted via **`[Copy SVG]`** is ungrouped (`Convert to Shape`), its chemical link is replaced by raw vector polygons—always keep an untouched original copy on a backup slide.
* **Verify Tags & Alternative Open Methods:** Right-click the shape ➔ **View Alt Text** to confirm `SMILES:`, `KET:`, or `MOLSTAR:` is intact. You can always select the shape and click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) or **`[3D Structure (Mol*)]`** (3D). If the `[SMILES]` tab itself is missing, check [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab).

## 3. Presenting on Conference Podium PCs or macOS {#sharing}

* **Zero Add-In Required on Podium Machines:** Inserted 2D PNGs and 3D `.glb` models are native PowerPoint shapes embedded inside `.pptx`—your slides, 3D **Turntable / Swing** animations, and **Morph** transitions play reliably on **conference PCs and macOS** without installing `func_SMILES2IMG` ([Document Sharing](/guide/troubleshooting/document-sharing)).
* **Exporting to PDF Handouts:** PDF flattens `.glb` objects at their current resting camera angle—verify every 3D slide rests on your intended pocket view before saving as PDF.

## Related Guides

- [**Dark Slides & Grouped Shapes**](/powerpoint/troubleshooting/dark-slides) — Fix white rectangles, invisible dark bonds, and grouped SAR cards skipped by `Batch Style`
- [**3. 360° 3D Protein Slides (`Mol*` & Morph)**](/powerpoint/3d-proteins) — Live 3D Morph zoom-ins and lightweight `.glb` tips
- [**Offline PDB Files**](/editor/troubleshooting/offline-pdb) · [**Document Sharing**](/guide/troubleshooting/document-sharing) — Offline crystal structures and cross-platform presentation checks
