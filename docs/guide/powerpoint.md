# Microsoft PowerPoint Guide

Learn how to place high-resolution 2D chemical structures and reaction schemes onto presentation slides, switch entire decks to Dark Mode in one click, and present 360°-rotatable 3D protein models.

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 3 Key Advantages for PowerPoint Presentations
1. **Dark Mode & Transparent Backgrounds:** No more ugly white rectangular boxes on dark slides—export with alpha-transparent backgrounds or crisp white bonds in one click.
2. **Double-Click Editing Preserves Slide Layout:** **Double-clicking** a molecule on your slide updates it in place while keeping its exact slide position and size.
3. **Live 360° 3D Proteins on Slides:** Instead of static screenshots, insert real 3D protein models (`.glb`) that you can rotate and zoom with your mouse right during your talk!
:::

---

## 1. Place Structures & Reaction Schemes on Slides (3 Ways)

### Way ① Draw by Hand (`Insert / Edit Structure`)
1. On any PowerPoint slide, click **`[SMILES] ➔ [Insert / Edit Structure]`** on the ribbon (or right-click an empty slide area ➔ **`Chemical Structure ➔ Insert / Edit Structure (Ketcher)...`**).
2. Sketch your molecule or multi-step synthesis pathway (arrows, reagents, catalysts) in the **Ketcher 2D Editor**.
   * *Tip: For dark or colored slides, toggle **`[Background: Transparent]`** in the top bar of Ketcher.*
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to place the rendered structure right at the center of your slide!

### Way ② Convert a Text Box or Paste from Clipboard in 0.1s (`Paste Structure`)
* **Convert a Text Box:** Select any text box containing a SMILES string and click **`[SMILES] ➔ [Paste Structure]`** to turn it into a crisp molecule shape immediately.
* **Paste from Clipboard:** Copy any SMILES string, then click **`[SMILES] ➔ [Paste Structure]`** (or right-click a blank area of your slide ➔ **`Chemical Structure ➔ Paste Structure from Clipboard`**).

---

## 2. Match Dark Mode or Print Handouts in 1 Click (`Batch Style`)

Switching your slide deck from a light theme to a dark theme? Click **`[SMILES] ➔ [Batch Style]`** on the ribbon to restyle every chemical structure across your presentation (or selection) at once:

| Style Preset | Slide Visual Result | Best For |
| :--- | :--- | :--- |
| **Standard Color** | IUPAC element colors + **transparent background** | Light-colored or gradient presentation themes |
| **Dark Mode** | **Crisp white bonds** + transparent background | **Dark presentation themes** and auditorium OLED screens |
| **Publication B&W** | High-contrast **monochrome line art (ACS style)** | Printed seminar handouts and lecture notes |
| **White Background** | Element colors + **solid white background** | Placing structures over busy photos or textured backgrounds |

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

