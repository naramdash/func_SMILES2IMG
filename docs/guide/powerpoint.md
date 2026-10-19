# PowerPoint: Dark Mode Slides & 360° 3D Proteins

Place transparent 2D chemical structures onto dark presentation slides without ugly white boxes, restyle entire slide decks in one click, and rotate 3D protein models live during your talk.

<p align="center">
  <img src="/powerpoint-showcase.png" alt="func_SMILES2IMG PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🎯 How to Start in PowerPoint (3 Steps: ① ➔ ② ➔ ③)

Unlike Excel, PowerPoint uses the top **`[SMILES]`** ribbon tab or keyboard shortcut <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>:

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **① Click `[SMILES]` ➔ `[Insert / Edit Structure]`** on the top ribbon (or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
2. **② Sketch & Insert (`Ctrl + Enter`):** Draw your molecule or synthesis pathway in Ketcher (toggle **`[Background: Transparent]`** in the top bar for dark slides!) and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
   * *Have a SMILES string on your clipboard?* Click **`[Paste Structure]`** to place a crisp molecule shape on your slide in 0.1s!
3. **③ Double-Click Anytime to Edit:** **Double-click** any slide molecule to update it in place while keeping its exact slide coordinates and size.

---

## 2. Dark Mode (White Ink) vs. Standard Color in 1 Click (`Batch Style`)

Stop pasting white-background screenshots onto dark conference slides! Click **`[SMILES] ➔ [Batch Style]`** on the ribbon to switch all structures across your slide deck between **Standard Color**, **Dark Mode (White Ink + Transparent BG)**, **Publication B&W**, or **White Background**:

<BatchStylePreview />

---

## 3. Present 360°-Rotatable 3D Proteins (`Mol*`)

Explain binding pockets and protein conformations interactively instead of pasting multiple static angles:

1. Click **`[SMILES] ➔ [3D Structure (Mol*)]`** on the ribbon.
2. Enter a 4-letter PDB ID (e.g., `1CRN`, `4HHB`, `6LU7`) to download, or open a local structure file.
3. Customize the representation (Cartoon, Molecular Surface, Ball & Stick) and click **`[Insert 3D Model (.glb)]`**.
4. Right on your slide, drag the central 3D handle with your mouse to **rotate and zoom the protein 360°** in front of your audience!  
   *(You can also choose **`[Insert 2D Picture (.png)]`** with a transparent background whenever you prefer a flat high-res image.)*

---

## 4. Paste Bullet-Point Summaries & Copy Vector SVG for Posters

**Right-click** any molecule shape on your slide and select **`Chemical Structure`**:

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **Instant Slide Bullet Points:** Click **`Copy Molecule Info ➔ Copy as Multi-Line List`** and paste into any PowerPoint text box to get a ready-made bulleted list of Name, Formula, and Molecular Weight!
* **Lossless Vectors for Conference Posters (`Copy SVG`):** Click **`[Copy SVG]`** on the ribbon (or right-click ➔ **`Copy Vector SVG`**) to copy a scalable vector graphic that stays razor-sharp even on large-format conference posters or in Adobe Illustrator.

---

## 5. Top `[SMILES]` Ribbon Overview

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ `Insert / Edit Structure`:** Open the 2D molecule & reaction sketcher (or edit the selected slide shape)
* **🧬 `3D Structure (Mol*)`:** Open the interactive 3D protein viewer
* **📋 `Paste Structure`:** Convert a selected SMILES text box or clipboard data into a slide shape in 0.1s
* **🖼️ `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** Copy chemical identifiers or lossless SVG vectors for conference posters
* **🔬 `Molecule Info`:** Copy multi-line property lists or open the full property summary window
* **🎨 `Batch Style`:** Convert all structures across your slide deck to Dark Mode, Color, B&W, or White Background

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../guide/molstar" class="s2i-card">
    <div class="s2i-card-title">🧬 Mol* 3D Protein & Crystal Viewer</div>
    <p class="s2i-card-desc">Try 4 sample PDB codes (1CRN, 4HHB, 6LU7, 1BNA) and capture 3D binding pockets</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 Publishing, SVG & Property Tables</div>
    <p class="s2i-card-desc">Copy lossless SVG vectors for conference posters and test Dark Mode batch styling</p>
  </a>
</div>

