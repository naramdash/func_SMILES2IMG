# PowerPoint Use Cases

Explore 5 practical ways **`func_SMILES2IMG`** elevates chemistry slides and posters—from dark-mode conference talks with live 3D protein motion to weekly 4-card SAR decks, meter-wide A0 posters, Q&A speaker notes, and animated mechanisms.

| Use Case | What You Can Do | Key Tool |
| :--- | :--- | :--- |
| [**1. Conference Talk & 3D Motion**](#conference-talk) | Match dark templates cleanly and animate native `.glb` 3D protein fly-ins via Morph | `[Batch Style]`, `[3D Structure (Mol*)]` |
| [**2. Weekly Lab Meeting SAR Grid**](#sar-review) | Convert candidate TextBoxes in `0.1s` and keep 4-card layouts aligned between tweaks | `[Paste Structure]`, Double-Click |
| [**3. A0 Posters & Print Handouts**](#a0-posters) | Paste scalable `SVG` paths onto A0 canvases and switch a deck copy to ACS B&W in 1s | `[Copy SVG]`, `[Publication B&W]` |
| [**4. Mechanism Seminar & Q&A Notes**](#qa-seminar) | Move `KET` schemes from Word/Excel onto dark slides and preload Q&A Speaker Notes | `[Copy KET]`, `Multi-Line List` |
| [**5. Step-by-Step SVG Animation**](#svg-step-animation) | Split pasted SVG drawings into native PowerPoint shapes for custom highlight & click build | `[Copy SVG]` ➔ `Convert to Shape` |

---

## 1. Conference Talk & 3D Motion {#conference-talk}

Replace flat PyMOL screenshots and codec-risky MP4 videos with zero-white-box dark-mode structures and native **3D `.glb` proteins** that spin continuously on slide, zoom smoothly into binding pockets via **Morph**, and run natively on any podium PC or Mac.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

1. **Apply Dark Mode Across the Deck:** Insert 2D molecules or schemes, click empty slide canvas, and run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`**.
2. **Embed a Native 3D `.glb` Protein:** Click **`[SMILES]` ➔ `[3D Structure (Mol*)]`**, fetch a PDB ID (`1CRN`, `6LU7`) or open a local `.pdb`/`.cif` file, style the binding pocket, and insert as **`3D Model (GLB)`**.
3. **Animate Motion Natively in PowerPoint:**
   - **Continuous Spin (`Animations`):** Select the `.glb` protein and apply **`Turntable`** (smooth 360° spin), **`Swing`** (gentle front rock preserving pocket view), or **`Arrive`**.
   - **Pocket Fly-In (`Transitions ➔ Morph`):** Duplicate the slide, rotate and zoom the `.glb` model on Slide B tightly onto the binding pocket, and enable **Morph** to animate the camera smoothly between slides.
4. **Zero-Risk Podium Playback:** Embedded `.png`, `.glb`, 3D animations, and Morph transitions play natively on conference PCs and macOS without installing `func_SMILES2IMG`.

* **Next:** [**3. 360° 3D Protein Slides (`Mol*` & Morph)**](/powerpoint/3d-proteins) · [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) · Playbook **Slides 1 & 6**

---

## 2. Weekly Lab Meeting SAR Grid {#sar-review}

Assemble aligned 4–8 analog comparison slides in minutes and tweak substituents for next week's group meeting with a quick double-click instead of rebuilding textboxes and picture alignment from scratch.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. **Convert TextBoxes In-Place (`0.1s`):** Lay out 4 candidate cards with TextBoxes containing SMILES or compound names (`Aspirin`, `Caffeine`, `Paracetamol`), select each TextBox border, and click **`[SMILES]` ➔ `[Paste Structure]`**.
2. **Unify Visual Style in 1s:** Run **`[SMILES]` ➔ `[Batch Style]`** (`Dark Mode` or `Standard Color`; ungroup cards first if grouped with `Ctrl+G`).
3. **Add Verified ADME Footnotes:** Click each candidate and choose **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** to paste Formula, MW, LogP, TPSA, and `Fsp3` beneath the card.
4. **Update Next Week in Seconds:** **Double-click** any card's structure in Normal view to modify R-groups in Ketcher—its card center and footprint stay locked automatically.

* **Next:** [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) · [**1. 2D Structures & Reaction Slides**](/powerpoint/slides-structures) · Playbook **Slides 3 & 4**

---

## 3. A0 Posters & Print Handouts {#a0-posters}

Avoid jagged bitmaps when printing **90 × 120 cm A0 conference posters** at 300 DPI, and generate clean black-and-white PDF handouts from the exact same `.pptx` file in one click.

<FeatureMedia id="powerpoint-slide-batch-style" />

1. **Paste Infinite-Resolution Vector `SVG`:** Select any 2D molecule or reaction scheme (or highlight SMILES/name text inside a TextBox), click **`[SMILES]` ➔ `[Copy SVG]`**, and press <kbd>Ctrl</kbd> + <kbd>V</kbd> onto your PowerPoint A0 poster (Office 2019+/365) or into **Adobe Illustrator / Inkscape**.
2. **Export Monochrome Handouts in 1s:** Save a copy of your dark deck, click empty canvas so nothing is selected, and run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** before exporting to PDF.

* **Next:** [**4. A0 Posters (SVG), Q&A, Podium Sharing**](/powerpoint/publishing-qa) · [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) · Playbook **Slides 7 & 8**

---

## 4. Mechanism Seminar & Q&A Notes {#qa-seminar}

Bring `KET` reaction schemes directly over from your Word paper onto dark seminar slides, and keep exact mass (`Da`), `Fsp3`, LogP, and TPSA answers right in front of you during audience Q&A.

<FeatureMedia id="powerpoint-slide-reaction" />

1. **Reuse Paper Schemes in 3s:** Select any `KET` reaction scheme in Word or Excel and click **`[SMILES]` ➔ `[Copy KET]`**. Switch to PowerPoint, click **`[Paste Structure]`**, and run **`[Batch Style]` ➔ `[Dark Mode]`** *(apply PowerPoint animations after finishing chemical edits)*.
2. **Preload Q&A Speaker Notes:** Click **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** on key compounds and paste into **Speaker Notes** so exact values appear privately in Presenter View.
3. **Inspect All 12 Descriptors Live:** After finishing full-screen slides (<kbd>Esc</kbd>), click any structure and open **`[Molecule Info]` ➔ `[Detailed Molecule Info Window...]`** to show calculated properties side by side.

::: tip Cross-App Suite Handoff
Screen libraries in [**Excel Use Cases**](/excel/use-cases) ➔ draft manuscripts in [**Word Use Cases**](/word/use-cases) ➔ bring those exact `KET` schemes and `SMILES` directly into your PowerPoint slides and posters.
:::

* **Next:** [**1. 2D Structures & Reaction Slides**](/powerpoint/slides-structures) · [**4. A0 Posters (SVG), Q&A, Podium Sharing**](/powerpoint/publishing-qa) · Playbook **Slides 5 & 7**

---

## 5. Step-by-Step SVG Animation {#svg-step-animation}

Reveal multi-step reaction mechanisms stage by stage on click or recolor key reactive bonds directly inside PowerPoint 2019–365 using editable native vector shapes.

1. **Finalize Chemistry First:** Finish all bonds, arrows, and labels in **Ketcher 2D**, then select the slide shape (or highlight a SMILES string).
2. **Paste as SVG Graphic:** Click **`[SMILES]` ➔ `[Copy SVG]`** and press <kbd>Ctrl</kbd> + <kbd>V</kbd> onto the slide.
3. **Convert SVG Into Native Shapes:** Right-click the pasted SVG ➔ **Group ➔ Ungroup** (or **Graphics Format ➔ Convert to Shape**) to split every bond line, arrow, and atom letter into individual vector objects.
4. **Recolor & Animate Stages:** Color key bonds or leaving groups using **Shape Outline / Fill**, re-group (<kbd>Ctrl</kbd> + <kbd>G</kbd>) Reactants, Arrow/Reagents, and Product separately, and apply **Appear** or **Wipe (From Left)** animations in sequence.

::: tip Keep One Editable Backup Shape
Once converted into raw polygons, a shape is no longer linked to Ketcher (`Alt Text` / `Shape.Tags` are cleared). Keep one original Ketcher-editable copy on a hidden backup slide in case you revise the scheme later.
:::

* **Next:** [**4. A0 Posters (SVG), Q&A, Podium Sharing**](/powerpoint/publishing-qa) · [**Style & Export Reference**](/editor/publishing)

---

## Next Steps

- [**Practice in the 8-Slide PowerPoint Playbook (.pptx)**](/powerpoint/playbook) — Test TextBox `0.1s` conversion, 4-card dark grids, `KET` mechanisms, and `1CRN` Morph transitions in PowerPoint.
- [**10-Second Install Guide**](/guide/install) — Install across PowerPoint, Word, and Excel with one PowerShell command.
