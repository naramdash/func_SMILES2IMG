# SMILES for Microsoft PowerPoint

**SMILES for Microsoft PowerPoint** brings effortless chemical presentations to life with one-click slide insertion, dark theme transparent backgrounds, and seamless in-place editing.

---

## 🌟 Visual Feature Walkthrough

Crafting scientific slides, seminar presentations, and conference decks with chemical accuracy has never been easier:

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip Key Advantages in PowerPoint
* **Dark Theme Optimization:** Switch between crisp white backgrounds and alpha-transparent PNGs at any time so chemical structures blend seamlessly into modern dark slide decks without white rectangular halos.
* **Exact Geometric Preservation:** When double-clicking to edit, the new structure retains the exact slide coordinates (`Left`, `Top`), dimensions, and center anchor point.
* **Lossless Slide Portability:** Chemical metadata is stored directly in PowerPoint shape tags (`Shape.Tags['KET']` and `Shape.Tags['SMILES']`), allowing seamless cross-machine editing.
:::

---

## 📑 Dedicated Office Ribbon: `[SMILES]` Tab

The custom `[SMILES]` tab on PowerPoint's ribbon provides one-click slide insertion and clipboard parsing:

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ Insert Structure (`Ctrl+Shift+K`):** Launches Ketcher 2D editor. Draw molecules or multi-step reactions, then click **Apply** (<kbd>Ctrl</kbd> + <kbd>Enter</kbd>) to place them at the center of the active slide.
2. **🧬 3D Structure (Mol*):** Opens the embedded Mol* 3D viewer. Download PDB structures (e.g. `1CRN`, `4HHB`) or open local files, then insert as an **interactive 3D model (.glb)** or high-resolution 2D image. See the [Mol* 3D Viewer Guide](/guide/molstar) for full details.
3. **⚡ Quick Insert:**
   - If a text box or shape containing SMILES text is selected on the slide, converts the text frame into a high-resolution molecule shape in 0.1s.
   - If no shape is selected, automatically reads chemical text from the clipboard and places it onto the slide.
4. **📋 Paste Structure:** Reads clipboard SMILES/KET and embeds it as a vector-crisp slide image.
5. **🖼️ Copy SVG *(New in v1.7.0)*:** Renders the selected molecule shape as a scalable vector graphic (`SVG`) and copies it to the Windows Clipboard.
6. **🎨 Batch Style *(New in v1.7.0)*:** Converts all chemical structures across the active presentation to a unified style preset (**Standard Color (Transparent)**, **Publication B&W (ACS Monochrome)**, **White Background (Color)**, or **Dark Mode (White Ink)**) in one click.
7. **🔍 Settings & Diagnostics:** Opens engine diagnostic logs and troubleshooting metrics.

---

## 🖱️ Redesigned Context Menu: `Chemical Structure`

Right-click any chemical shape or blank presentation slide area:

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Case A: Right-Clicking an Existing Chemical Shape or 3D Model
* **✏️ Edit Structure in Ketcher:** Reopens Ketcher with the current molecule or reaction loaded.
* **🧬 Edit 3D Structure (Mol*):** Reopens the Mol* 3D viewer in the exact camera orientation and representation style.
* **📋 Copy SMILES:** Copies clean SMILES notation to the clipboard.
* **📋 Copy InChI:** Copies standard IUPAC InChI string instantly.
* **📋 Copy Canvas Data (KET):** Copies Ketcher native JSON representation.
* **🖼️ Copy Vector Graphic (SVG):** Renders the structure as a scalable vector graphic (`SVG`) and copies it to the clipboard for Adobe Illustrator or Inkscape.
* **🔬 Copy Molecule Info (Submenu):**
  * **Substance Name:** Copies substance name (PubChem / IUPAC).
  * **Chemical Formula:** Copies standard Hill-notation formula (e.g. `C9H8O4`).
  * **Molecular Weight (MW):** Copies IUPAC molecular weight in g/mol (e.g. `180.16 g/mol`).
  * **Copy SMILES / InChI / InChIKey:** Individual one-click copy options.
  * **Copy All Summary (Text):** Copies compact single-line summary to clipboard.
  * **Copy as Multi-Line List:** Copies formatted multi-line property list for slide bullet points.
  * **Copy as Excel Row (TSV Columns) / Markdown Table:** (Available from the Ribbon `[Molecule Info]` menu and Detailed Window) Exports properties as TSV/HTML table rows or Markdown tables.
  * **Detailed Molecule Info Window...:** Opens the full chemical properties dialog (including **LogP / TPSA**, **Lipinski Rule of 5**, and split-button multi-format exporter).
  > ⚡ **Non-Blocking Toast:** Each copy action displays a floating pill badge (`✓ Copied {Item}: {Value}`) near your mouse for 1.4s without stealing PowerPoint's focus.

### Case B: Right-Clicking Blank Slide Canvas
* **✏️ Insert Structure in Ketcher:** Launches Ketcher to draw a new molecule onto the slide.
* **🧬 3D Structure (Mol*)...:** Launches Mol* to insert a 3D macromolecular model or image.
* **📋 Paste Structure:** Reads chemical notation from clipboard and creates a slide shape.

---

## 🧪 Ketcher 2D Molecule & Reaction Editor

Double-click any molecule shape on your slide or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> to launch Ketcher (with `<0.05s` warm WebView2 instance pooling for instant reopening):

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **🌓 Dark Theme & Transparent Backgrounds:** Toggle between **Background: White** and **Background: Transparent** in the top status bar. Transparent mode exports clean alpha-channel PNGs, ensuring no awkward white rectangular borders appear around your chemical bonds on dark backgrounds.
* **Live Molecular Property Banner:** Live updates for Formula, MW, and Exact Mass.
* **Hotkeys:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to update the slide instantly, <kbd>Esc</kbd> to cancel.

---

## 🔬 Multi-Step Reactions & Synthesis Schemes

Present complex chemical mechanisms, catalytic cycles, and organic synthesis pathways directly on your slides:

* **Reaction Schemes Preserved:** Full reaction schemes with arrows, catalysts, and reagents are stored losslessly inside `Shape.Tags['KET']`.
* **In-Place Re-editing:** Double-clicking the slide shape reopens the reaction canvas in Ketcher, preserving all arrows and conditions for quick layout adjustments before presentations.

---

## 🔍 Built-In Diagnostics Console

Click **[SMILES] ➔ [Settings & Diagnostics]** to view system health and engine logs:

* Confirms PowerPoint COM Add-in registration and Indigo engine readiness.
* Verifies offline Ketcher assets and rasterization performance metrics.
