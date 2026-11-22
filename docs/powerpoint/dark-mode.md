# 2. Dark Mode, Candidate Cards, Deck Restyle

Render crisp white-ink chemical structures over dark navy, slate, or gradient slides with zero white rectangles—and switch an entire multi-slide presentation between Dark Mode, bright seminar templates, and monochrome handouts in one click.

## 1. Two Ways to Eliminate White Boxes on Dark Slides

Standard screenshots leave blinding white boxes on dark conference themes. With `func_SMILES2IMG`, both background and ink adapt cleanly:

<FeatureMedia id="powerpoint-slide-dark-mode" />

* **Method A — Before Inserting (Ketcher / Mol\*):** Switch Ketcher's top bar to **`[Background: Transparent]`** (or Mol\*'s **`BG: Transparent`**) before pressing <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
* **Method B — After Inserting (1-Click `Batch Style`):** Select one or more shapes (or leave nothing selected to convert the **entire deck**) and click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`**.

## 2. 4-Card Lead Candidate SAR Grid Recipe

Build polished Hit-to-Lead and SAR comparison slides where structures, potency tags, and ADME summaries line up pixel-for-pixel:

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. Set a dark slide background (`#0F172A` slate or `#1E293B` navy) and draw 4 equal rounded-rectangle cards side by side.
2. Add each candidate via **`[Paste Structure]`** (select a typed SMILES/name TextBox) or **`[Insert / Edit Structure]`**, then align the structures horizontally.
3. Click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`** so all 4 cards share identical crisp white line weight and transparent backgrounds.
4. Click each structure and choose **`[SMILES]` ➔ `[Molecule Info]` ➔ `[Copy as Multi-Line List]`** (or **`Copy Molecular Weight`** / **`Detailed Window`**) to paste verified Formula, MW, LogP, TPSA, and Lipinski metrics into the text area below.
5. **Tweak Substituents Without Re-Aligning:** **Double-click** any card's molecule anytime to modify R-groups in Ketcher—when applied, the structure stays centered inside its original card footprint.

## 3. Switch an Entire Presentation's Theme in 1 Second (`Batch Style`)

Reusing slides from a dark auditorium conference deck inside a white lab-meeting template or printed handout takes one click across every slide at once:

<FeatureMedia id="powerpoint-slide-batch-style" />

1. **Choose Your Scope:**
   * **Entire Presentation Deck:** Click empty slide canvas (nothing selected). Every 2D molecule and `KET` reaction across **all slides** updates automatically (`Mol*` 3D objects stay untouched).
   * **Selected Shapes Only:** Hold <kbd>Shift</kbd> and click specific structures to restyle only those candidates.
2. Click **`[SMILES]` ➔ `[Batch Style]`** and pick the matching preset (see [Style & Export Reference](/editor/publishing) for exact specs):
   * **`Dark Mode (White Ink)`** — Transparent BG + crisp white ink for dark themes
   * **`Standard Color (Transparent)`** — Transparent BG + IUPAC element colors for light templates
   * **`Publication B&W (ACS Monochrome)`** — Transparent BG + pure black ink for printed handouts
   * **`White Background (Color)`** — Solid white card + IUPAC colors over busy photo/gradient backgrounds

::: warning Grouped Shapes Are Skipped by `Batch Style`
PowerPoint `Batch Style` scans top-level slide shapes. If you **Grouped** (<kbd>Ctrl</kbd> + <kbd>G</kbd>) a molecule with its card text box, **Ungroup** (<kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>G</kbd>) before running **`[Batch Style]`**, then re-group afterward.
:::

## Related Guides

- [**1. 2D Structures & Reaction Slides**](/powerpoint/slides-structures) — 0.1s TextBox conversion, Ketcher reactions, and double-click editing
- [**4. A0 Posters (SVG), Q&A, Podium Sharing**](/powerpoint/publishing-qa) — Copy vector SVG for A0 posters and prepare speaker notes for Q&A
- [**Dark Slides & Grouped Shapes Troubleshooting**](/powerpoint/troubleshooting/dark-slides) — Fix white boxes, low-contrast bonds, or skipped grouped shapes
