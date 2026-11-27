# Dark Slides & Grouped Shapes

Fixes when a 2D chemical structure shows an opaque white rectangle or dark invisible bonds on dark slide themes, when grouped candidate cards are skipped by `Batch Style`, or when switching dark conference slides back to white handouts.

## 1. White Box Background or Invisible Dark Lines on Dark Slides {#dark-slides}

Standard copy-pasted screenshots leave opaque white boxes over dark navy/slate presentation templates, while transparent images with black bonds disappear over dark backgrounds.

* **Fix Existing Slide or Whole Deck Together:** Select the structure shape—or click empty slide background (nothing selected) to convert the **entire presentation deck**—and click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`**.
* **Insert Transparently From the Start:** Before pressing <kbd>Ctrl</kbd> + <kbd>Enter</kbd>, set Ketcher's top bar to **`[Background: Transparent]`** (or Mol\*'s **`BG: Transparent`**) so no opaque rectangle is generated.

<FeatureMedia id="powerpoint-slide-dark-mode" />

## 2. Grouped Cards Skipped by Batch Style {#grouped-shapes}

* **Why shapes get skipped:** `Batch Style` iterates over top-level slide shapes (`slide.Shapes`) and intentionally avoids mutating PowerPoint Groups (`Ctrl+G`) so card layouts and textboxes inside a group never detach or reorder.
* **Fix:** Select the grouped SAR card ➔ press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>G</kbd> (**Ungroup**) ➔ run **`[SMILES]` ➔ `[Batch Style]`** ➔ re-group (<kbd>Ctrl</kbd> + <kbd>G</kbd>) once styling is unified.

## 3. Switching Dark Slides Back to a Bright Template or Handout

You never need to redraw dark-mode structures when reusing conference slides inside a bright lab-meeting deck or printing monochrome handouts:

* **Light Templates:** Click empty slide canvas ➔ **`[SMILES]` ➔ `[Batch Style]` ➔ `[Standard Color (Transparent)]`** restores IUPAC heteroatom colors with alpha transparency.
* **Print Handouts / B&W PDFs:** Save a duplicate `.pptx` and run **`[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** to convert every slide structure to high-contrast black line art.

::: tip Animated Shapes Tip
If a structure has slide Entrance animations applied, wait until chemical edits are finished before adding animations, or re-apply the animation after `Batch Style` replaces the picture object.
:::

## Related Guides

- [**3D Models & Double-Click**](/powerpoint/troubleshooting/3d-editing) — Fix `.glb` size/compatibility limits, double-click editing, and podium PC checks
- [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) — Transparent 4-card candidate SAR grid layout and full `Batch Style` guide
- [**Ribbon & Context Menus**](/powerpoint/reference) — Complete PowerPoint `[SMILES]` ribbon map and 5 slide context menus
