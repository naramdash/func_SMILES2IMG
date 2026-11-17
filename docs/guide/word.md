# Word Manuscript Tutorial: Complete One Research Paper in 6 Chapters

> **Word Guide Style — Continuous 6-Chapter Manuscript Tutorial:** Follow a single research paper (*Synthesis and Structural Characterization of Aspirin Analogs*) from inserting your first inline figure in Chapter 1 all the way to pre-submission ACS Black & White conversion in Chapter 6. Need just one quick feature? Jump straight via the **Task Index** below.

<FeatureMedia id="word-overview" />

## Quick Task Index (Jump Directly by Task)

| What You Want to Do in Word | Where to Jump | Shortcut / Ribbon Action |
| :--- | :---: | :--- |
| Sketch a new 2D structure inline in prose | [Chapter 1](#chapter-1-insert-a-2d-structure-into-introduction-prose) | <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd> or `[Insert / Edit Structure]` |
| Turn typed SMILES text into a figure in 0.1s | [Chapter 2](#chapter-2-convert-inline-smiles-text-in-0-1s-paste-structure) | Highlight SMILES ➔ `[Paste Structure]` |
| Fill a Word table with structures & MW/Formula | [Chapter 3](#chapter-3-build-a-compound-summary-table-molecule-info) | Right-click ➔ `Copy Molecule Info` |
| Insert a multi-step reaction scheme | [Chapter 4](#chapter-4-insert-multi-step-reaction-schemes-lossless-ket) | Ketcher reaction arrows (`➔`) ➔ Apply |
| Edit any existing picture or add a 3D protein | [Chapter 5](#chapter-5-double-click-edits-3d-protein-figures-mol) | **Double-click** figure or `[3D Structure (Mol*)]` |
| Convert the whole manuscript to ACS B&amp;W / SVG | [Chapter 6](#chapter-6-pre-submission-acs-b-w-polish-vector-svg) | `[Batch Style]` &amp; `[Copy SVG]` |

---

## Chapter 1: Insert a 2D Structure into Introduction Prose

In Chapter 1 of your manuscript, you introduce Aspirin directly inside a paragraph—no screenshot cropping or external ChemDraw copy-paste needed.

<FeatureMedia id="word-ch1-insert-ketcher" />

1. Place your cursor in the Introduction paragraph where the structure belongs.
2. Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
3. Draw Aspirin in the offline **Ketcher 2D Editor** and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to embed it cleanly inline with your text.

---

## Chapter 2: Convert Inline SMILES Text in 0.1s (`Paste Structure`)

When drafting experimental sections or pasting SMILES from electronic lab notebooks, highlight any SMILES string right inside your Word paragraph and replace it in place.

<FeatureMedia id="word-ch2-quick-insert" />

1. Select the SMILES text in your Word paragraph (e.g. `Cn1cnc2c1c(=O)n(C)c(=O)n2C`).
2. Click **`[SMILES]` ➔ `[Paste Structure]`** on the ribbon (or right-click ➔ **`Chemical Structure ➔ Paste Structure`**).
3. Your highlighted text turns into a crisp, aspect-ratio-locked inline picture immediately. *(No text selected? It converts the SMILES/KET currently on your clipboard!)*

---

## Chapter 3: Build a Compound Summary Table & `Molecule Info`

In Chapter 3, you organize analogs inside a Word table and populate chemical formulas, molecular weights, and LogP without manual calculation.

<FeatureMedia id="word-ch3-compound-table" />

1. Place your cursor inside a Word table cell and insert each compound via **`[Insert / Edit Structure]`** or **`[Paste Structure]`**.
2. Select any inserted structure and click **`[SMILES] ➔ [Molecule Info]`** (or right-click ➔ **`Chemical Structure ➔ Copy Molecule Info`**).
3. Choose **`Copy as Excel Row (TSV Columns)`** to paste 17 tab-separated descriptors directly across Word table cells—or open **`Detailed Molecule Info Window...`** to copy individual values, Markdown tables, or JSON.

---

## Chapter 4: Insert Multi-Step Reaction Schemes (`Lossless KET`)

Chapter 4 adds **Scheme 1** showing salicylic acid + acetic anhydride reacting to give Aspirin, preserving arrows and reagent labels inside Word `AlternativeText` metadata.

<FeatureMedia id="word-ch4-reaction-scheme" />

* Draw your reaction arrows (`➔`), `+` signs, and catalyst annotations in Ketcher and click **Apply**.
* Even after saving, closing, and sharing the `.docx` with co-authors, all atomic coordinates and arrow geometry remain embedded inside the picture so **double-clicking** restores the exact reaction scheme.

---

## Chapter 5: Double-Click Edits & 3D Protein Figures (`Mol*`)

Need to adjust a substituent inside a table cell—or show how your drug candidate binds a macromolecular target?

<FeatureMedia id="word-ch5-edit-and-molstar" />

1. **Double-click any 2D picture** inside prose or a Word table to edit it in place—Word preserves the exact column widths and margins (`LockAspectRatio = msoTrue`).
2. To add a 3D protein figure, click **`[SMILES] ➔ [3D Structure (Mol*)]`**, type a 4-letter PDB code like `1CRN` or `6LU7`, style it (Cartoon / Surface), and insert either an interactive **3D Model (`.glb`)** or a high-DPI **2D Figure (`.png`)**.

---

## Chapter 6: Pre-Submission ACS B&W Polish & Vector SVG

Before submitting your manuscript to ACS, Wiley, Elsevier, or RSC, unify every figure across all pages and tables in one click.

<FeatureMedia id="word-ch6-publication-finish" />

1. Click **`[SMILES] ➔ [Batch Style] ➔ [Publication B&W]`** to restyle every structure in your selection—or the **entire manuscript**—into high-contrast monochrome line art. *(Use **`White Background`** if placing structures inside shaded table rows.)*
2. Select any key scheme and click **`[Copy SVG]`** to paste editable lossless paths directly into Adobe Illustrator or Inkscape.

---

## Word Top `[SMILES]` Ribbon & Right-Click Menu Reference

<p align="center">
  <img src="/word-context-menu-guide.png" alt="Live 3-Level Chemical Structure Context Menu in Microsoft Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **`Insert / Edit Structure`** (`Ctrl+Shift+K`), **`3D Structure (Mol*)`**, **`Paste Structure`**
* **`Copy SMILES`**, **`Copy InChI`**, **`Copy KET`**, **`Copy SVG`**
* **`Molecule Info`**, **`Batch Style`**, **`About func_SMILES2IMG`**

<div class="s2i-card-grid">
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 Publishing, SVG &amp; 5 Property Export Formats</div>
    <p class="s2i-card-desc">Compare all 4 Batch Style presets interactively and inspect all 5 copy formats</p>
  </a>
  <a href="../guide/powerpoint" class="s2i-card">
    <div class="s2i-card-title">📙 Next: PowerPoint Visual Slide Gallery</div>
    <p class="s2i-card-desc">Turn your manuscript figures into dark-mode slides and 3D rotatable presentations</p>
  </a>
</div>
