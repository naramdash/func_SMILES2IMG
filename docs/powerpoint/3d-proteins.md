# 3. 360° 3D Protein Slides (`Mol*` & Morph)

Turn any PowerPoint slide into an interactive 3D structural biology showcase: fetch PDB entries or open private `.pdb`/`.cif` files offline, embed native **3D `.glb` models** you can rotate 360° live during Slide Show, and animate smooth camera fly-ins with **PowerPoint Morph**.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

## 1. Insert a Rotatable 3D Model or Transparent 2D Figure

1. Place yourself on the target slide.
   ::: tip Power Tip: Preload a PDB Code From a Selected TextBox
   If a slide TextBox already displays a 4-character PDB ID like `1CRN` or `6LU7` (or a local `.pdb`/`.cif` path), select it **before** clicking **`[3D Structure (Mol*)]`**—Mol\* loads that entry automatically on startup.
   :::
2. Click **`[SMILES]` ➔ `[3D Structure (Mol*)]`** on the PowerPoint ribbon (or right-click ➔ **`Chemical Structure` ➔ `3D Structure (Mol*)...`**).
3. Type a 4-character PDB ID (`1CRN`, `4HHB`, `6LU7`, `1BNA`) and click **`Fetch`**, or click **`Open File...`** to load local `.pdb`, `.cif`, `.bcif`, `.ent`, `.mol`, `.sdf`, `.molj`, or `.molx` structures offline.
4. Style **Cartoon**, **Surface**, or **Ball & Stick** ligand pockets, choose your output mode on the top toolbar, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Slide`**):

| Mol\* Output Mode | How It Works on Slides | Best Slide Use Case |
| :--- | :--- | :--- |
| **`3D Model (GLB)`** | Embeds a native PowerPoint 3D object you can drag to rotate 360° in Normal view and **during live Slide Show** | Interactive target & binding-pocket demos (PowerPoint 2019 / 2021 / 2024 / 365) |
| **`2D Picture`** | Embeds a high-DPI `.png` snapshot (pick **`BG: Transparent`** for dark or gradient slides) | Very large viral/ribosome complexes (keeps `.pptx` lightweight), PDF handouts, and PowerPoint 2016 |

5. **Camera-Preserving Re-Edit:** **Double-click** any inserted 3D model or 2D protein figure in Normal view to reopen Mol\* at the exact saved camera orientation, zoom, and style (`CustomXMLParts`).

## 2. Killer Talk Effects: PowerPoint `Morph` & Built-In 3D Animations

Because `Mol*` inserts a native **PowerPoint 3D Model (`.glb`)**, you can drive smooth protein motion directly with PowerPoint—without rendering MP4s or worrying about podium video codecs:

### A. 2-Slide Binding-Pocket Fly-In (`Transitions ➔ Morph`)
1. Insert a **`3D Model (GLB)`** on **Slide A** and rotate it to show the **overall protein Cartoon structure**.
2. Right-click Slide A in the left thumbnail pane ➔ **Duplicate Slide** to create **Slide B**.
3. On **Slide B**, drag and scroll the 3D model to zoom tightly into the **ligand binding pocket** (and move it left/right to leave space for callout labels).
4. Select **Slide B**, go to the PowerPoint **Transitions** tab, and choose **Morph**—during your talk, advancing from Slide A to Slide B smoothly rotates and zooms the 3D protein live on screen.

### B. On-Slide Motion (`Animations ➔ 3D Animations`)
Select any inserted `.glb` protein on a slide and open PowerPoint's **Animations** tab to apply native 3D effects:
* **`Turntable` (Continuous 360° Spin):** Keeps the protein slowly rotating in place while you explain overall domain architecture.
* **`Swing` (Pocket Tilt):** Gently rocks the model back and forth around a fixed angle so audience depth perception catches ligand-pocket geometry without hiding the active site behind the protein.
* **`Arrive` (Entrance Spin):** Smoothly rotates the macromolecule into view when entering a slide. Use **Effect Options** to adjust rotation axis, intensity, and repeat loop.

## 3. File Size & Podium PC / macOS Compatibility

* **Zero Add-In Needed at the Podium:** Inserted `.glb` models and `.png` figures are standard native PowerPoint shapes. Your 3D rotations and Morph animations play normally on **conference podium PCs and macOS** without `func_SMILES2IMG` installed.
* **Keep Huge Complexes Lightweight:** Small-to-medium proteins (`1CRN`, `4HHB`, `6LU7`) load instantaneously as `.glb` models. For massive multi-subunit ribosomes or cryo-EM capsids—or if emailing a strict size-capped deck—choose **`2D Picture`** with **`BG: Transparent`**.

## Related Guides

- [**Mol\* 3D Viewer Manual**](/editor/molstar) — Sample PDB codes (`1CRN`, `4HHB`, `6LU7`, `1BNA`) and representation controls
- [**3D Models & Double-Click Troubleshooting**](/powerpoint/troubleshooting/3d-editing#3d-models) — Office version checks for `.glb` and offline `.pdb` loading behind firewalls
- [**Offline PDB Files Troubleshooting**](/editor/troubleshooting/offline-pdb) — Load local structure files without internet access
