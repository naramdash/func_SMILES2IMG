# Microsoft Word Guide

Learn how to insert and edit high-resolution 2D chemical structures, reaction schemes, and 3D proteins inside Word manuscripts, journal articles, thesis chapters, and tables.

<p align="center">
  <img src="/word-showcase.png" alt="SMILES for Word Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 3 Key Advantages for Word Manuscripts
1. **Natural Inline Flow:** Structures are inserted as inline pictures that flow naturally within sentences and align cleanly inside table cells.
2. **Double-Click Editing Without Breaking Tables:** **Double-clicking** a structure inside a table updates it in place without messing up your column widths or row heights.
3. **Editable Across Computers:** The underlying chemical structure is saved inside the `.docx` file itself, so co-authors on other PCs can double-click to edit at any time.
:::

---

## 1. Insert Structures & Reactions into Text or Tables (3 Ways)

### Way ① Draw by Hand (`Ctrl + Shift + K`)
1. Place your cursor in any paragraph or table cell in Word.
2. Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES] ➔ [Insert Structure]`** on the ribbon).
3. Draw your molecule or multi-step reaction scheme (arrows, reagents, catalysts, temperatures) in the **Ketcher 2D Editor**.
4. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert the structure right at your cursor!

### Way ② Convert Typed SMILES Text in 0.1s (`Quick Insert`)
* Highlight any SMILES text typed in your document (e.g., `c1ccccc1`) and click **`[SMILES] ➔ [Quick Insert]`** to replace the text with a rendered structure picture on the spot.

### Way ③ Paste from Clipboard
* Copy a SMILES string from a paper, website, or ChemDraw, then click **`[SMILES] ➔ [Paste Structure]`** (or right-click ➔ **`Chemical Structure ➔ Paste Structure`**).

---

## 2. Convert Your Entire Manuscript to Journal B&W in 1 Click (`Batch Style`)

Before submitting to a journal, you don't need to redraw or re-export dozens of figures. Click **`[SMILES] ➔ [Batch Style]`** on the ribbon to unify every structure in your document (or selection) in one click:

| Style Preset | Visual Result | Best For |
| :--- | :--- | :--- |
| **Standard Color** | IUPAC element colors + **transparent background** | Color thesis chapters, digital PDFs, and internal reports |
| **Publication B&W** | High-contrast **monochrome line art (ACS style)** | Print journal submissions (ACS, Wiley, Elsevier, RSC) & patents |
| **White Background** | Element colors + **solid white background** | Shaded Word table rows or colored callout boxes |
| **Dark Mode** | **White bonds** + transparent background | Dark-themed covers and digital documents |

---

## 3. Fill Word Tables with Properties & Copy Vector SVG

**Right-click** any chemical structure picture in Word and select **`Chemical Structure`** (or use the top ribbon) to speed up writing your Experimental section:

<p align="center">
  <img src="/word-context-menu-guide.png" alt="Chemical Structure Context Menu in Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### 📋 Paste Properties Directly into Word Tables & Paragraphs
* **1-Click Property Copy:** Copy Substance Name, Formula (`C9H8O4`), Molecular Weight (`180.16 g/mol`), Exact Mass, or InChIKey individually.
* **Auto-Fill Word Table Rows (TSV):** Click **`Copy Molecule Info ➔ Copy as Excel Row (TSV Columns)`** and paste directly into a Word table—each property automatically populates its own table cell!
* **Detailed Molecule Info Window:** View all calculated properties (including LogP, TPSA, and Lipinski Rule of 5) in one window and export as a Markdown table or JSON.

### 🖼️ Copy Lossless Vector Graphics for Illustrator (`Copy SVG`)
* Click **`Copy Vector Graphic (SVG)`** to copy a true scalable vector graphic (`SVG`) to the clipboard, ready to paste directly into Adobe Illustrator, Inkscape, or ChemDraw without losing resolution.

---

## 4. Embed 3D Proteins in Manuscripts (`Mol*`)

1. Click **`[SMILES] ➔ [3D Structure (Mol*)]`** on the ribbon.
2. Enter a 4-letter PDB ID (e.g., `1CRN`, `4HHB`) to download, or open a local structure file (`.pdb`, `.cif`, `.sdf`).
3. Choose your visual style (Cartoon ribbon, Ball & Stick, Molecular Surface) and camera angle, then insert as:
   * **[3D Model (.glb)]:** An interactive 3D object you can rotate 360° with your mouse right inside Word.
   * **[2D Picture (.png)]:** A high-resolution figure for print publication and PDF export.
4. **Double-click** any inserted 3D model or picture at any time to reopen Mol\* at the exact same camera angle.

---

## 5. Top `[SMILES]` Ribbon Overview

<p align="center">
  <img src="/word-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ Insert Structure (`Ctrl+Shift+K`):** Open the 2D molecule & reaction sketcher
* **🧬 3D Structure (Mol\*):** Open the 3D protein & macromolecular viewer
* **⚡ Quick Insert:** Convert highlighted SMILES text (or clipboard) into a picture in 0.1s
* **📋 Paste Structure:** Insert copied SMILES/KET data as an inline picture
* **🖼️ Copy SVG:** Copy the selected structure as a lossless SVG vector
* **🎨 Batch Style:** Convert all structures in the document to Color, Publication B&W, White Background, or Dark Mode
* **🔬 Molecule Info:** Inspect and copy MW, Formula, LogP, Lipinski Ro5, and table rows

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 Publishing, SVG & Property Tables</div>
    <p class="s2i-card-desc">Compare ACS B&W batch presets interactively and explore 5 property export formats</p>
  </a>
  <a href="../guide/powerpoint" class="s2i-card">
    <div class="s2i-card-title">📽️ PowerPoint Presentations</div>
    <p class="s2i-card-desc">Create dark-mode slide structures and embed 360° rotatable 3D protein models</p>
  </a>
</div>

