# PowerPoint Visual Slide Gallery: Dark Mode Decks & 3D Proteins

> **PowerPoint Guide Style — Visual Slide Gallery:** Browse finished presentation slide patterns below, preview the live-recorded video and screenshot of each layout, and follow 3 quick steps to build the slide in your own deck.

<FeatureMedia id="powerpoint-overview" />

---

## Slide Pattern 1: 4-Card Lead Compound Screening Grid (Dark Theme)

Build conference-ready dark-slate screening cards with transparent chemical structures—zero ugly white bounding boxes.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. Set your slide background to dark navy/slate (`#0F172A`) and lay out compound cards.
2. Insert each compound via **`[SMILES] ➔ [Paste Structure]`** or **`[Insert / Edit Structure]`**.
3. Click **`[SMILES] ➔ [Batch Style] ➔ [Dark Mode]`** to convert all slide molecules into crisp white-ink transparent structures.

---

## Slide Pattern 2: Dark Mode White Ink with Transparent Canvas

Compare standard light styling against conference-projector Dark Mode (white carbon bonds, vivid heteroatoms, transparent alpha background).

<FeatureMedia id="powerpoint-slide-dark-mode" />

* **Shortcut inside Ketcher:** Toggle **`[Background: Transparent]`** in Ketcher's top bar before clicking **Apply to Slide**.
* **Batch Switch Later:** Select any set of slide pictures (or leave none selected to update the **entire deck**) and click **`[Batch Style] ➔ [Dark Mode]`**.

<BatchStylePreview />

---

## Slide Pattern 3: Multi-Step Synthetic Mechanism Slide

Place reaction pathways with arrows, `+` signs, and catalyst conditions cleanly on slides—with full KET coordinates stored inside `Shape.Tags["KET"]`.

<FeatureMedia id="powerpoint-slide-reaction" />

1. Click **`[SMILES] ➔ [Insert / Edit Structure]`** (<kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
2. Sketch reactants, reaction arrows (`➔`), and reagent text (`T`), then press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
3. **Double-click** the reaction graphic anytime on your slide to tweak arrows or conditions in place without moving card alignments.

---

## Slide Pattern 4: Interactive 360°-Rotatable 3D Protein Slide (`Mol*`)

Let your audience see binding pockets and secondary structure from every angle—drag to rotate 3D models live during your presentation.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

1. Click **`[SMILES] ➔ [3D Structure (Mol*)]`** on the PowerPoint ribbon.
2. Enter a 4-letter PDB ID (`1CRN`, `4HHB`, `6LU7`, `1BNA`) and click **`Fetch`**, or load a local `.pdb`/`.cif`/`.sdf` file.
3. Choose **`3D Model (GLB)`** (rotatable 360° on the slide) or **`2D Picture`** (`BG: Transparent`) and click **`Apply to Slide`**.

---

## Slide Pattern 5: Convert a Slide TextBox SMILES in 0.1s

Turn SMILES text boxes typed directly on a slide into centered vector-sharp structures in a single click.

<FeatureMedia id="powerpoint-slide-textbox-quick-insert" />

1. Select a PowerPoint TextBox containing a SMILES string (e.g. `CC(=O)Oc1ccccc1C(=O)O`) or compound name (`Aspirin`).
2. Click **`[SMILES] ➔ [Paste Structure]`** on the ribbon.
3. PowerPoint places the rendered molecule right on your slide with locked aspect ratio (`msoTrue`).

---

## Slide Pattern 6: Live Q&A Property Lookup (`Molecule Info`)

When a reviewer or audience member asks for molecular weight, exact mass, or Lipinski compliance during Q&A, inspect and copy all 17 descriptors straight from the slide shape.

<FeatureMedia id="powerpoint-slide-live-qa-info" />

1. Select any molecule shape on your slide and click **`[SMILES] ➔ [Molecule Info]`** (or right-click ➔ **`Chemical Structure ➔ Copy Molecule Info`**).
2. Choose **`Copy as Multi-Line List`** to paste ready-made bulleted properties straight into a slide text box, or open **`Detailed Molecule Info Window...`**.

---

## Slide Pattern 7: One-Click Slide Deck Theme Switcher (`Batch Style`)

Moving figures from a dark talk template onto a white handout deck? Switch the whole deck in one click.

<FeatureMedia id="powerpoint-slide-batch-style" />

1. Leave all shapes unselected to transform the **entire presentation**, or select specific shapes on one slide.
2. Open **`[SMILES] ➔ [Batch Style]`** and switch between **`Standard Color`**, **`Publication B&W`**, **`White Background`**, and **`Dark Mode`**.

---

## PowerPoint Right-Click Context Menu & Poster SVG

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Live 3-Level Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* Right-click any slide molecule ➔ **`Chemical Structure ➔ Copy Vector SVG`** to paste scalable vector graphics directly into A0 conference posters or Adobe Illustrator.
* Click **`[SMILES] ➔ [About func_SMILES2IMG]`** to check version info and updates.

<div class="s2i-card-grid">
  <a href="../guide/molstar" class="s2i-card">
    <div class="s2i-card-title">🧬 Mol* 3D Protein Viewer Deep Dive</div>
    <p class="s2i-card-desc">Explore 4 sample PDB structures and double-click camera angle restoration</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">📐 Publishing, Vector SVG &amp; Batch Style</div>
    <p class="s2i-card-desc">Copy poster-ready SVG paths and explore all 5 property export formats</p>
  </a>
</div>
