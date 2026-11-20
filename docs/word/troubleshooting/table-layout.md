# Table Layout & Styling

Fixes for structure sizing inside Word compound summary tables and applying 1-click **`Batch Style`** across a manuscript.

## 1. Keeping Compound Table Column Widths Intact

* When you double-click and replace an inline structure inside a Word table cell, `func_SMILES2IMG` locks aspect ratio (`LockAspectRatio`) and replaces the picture in place to match the existing layout.
* **Best Practice for Multi-Row Tables:** Right-click the table ➔ **AutoFit** ➔ **Fixed Column Width** so Word itself does not auto-shrink neighboring text columns when inserting wider reaction schemes.

## 2. `Batch Style` Only Changed One Structure (or Changed the Whole Document)

`Batch Style` uses your current Word selection to decide its scope:

* **Selection Active:** If you highlight specific paragraphs or table rows, **`[SMILES]` ➔ `[Batch Style]`** restyles **only the structures inside that selection**.
* **No Selection (Cursor Only):** If nothing is highlighted, `Batch Style` converts **every chemical structure across the entire Word document** in one click (ideal right before journal submission using **`Publication B&W`**).

## 3. Structures Look Blurry in Print or Large Posters

Word embeds crisp high-DPI PNGs that print cleanly on standard journal pages. For vector-level editing in Adobe Illustrator or Inkscape, select the structure and click **`[SMILES]` ➔ `[Copy SVG]`** (or right-click ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**).

## Related Guides

- [**Double-Click Editing**](/word/troubleshooting/double-click) — Verify Alt Text metadata and Word COM Add-in status if double-click fails.
- [**Word 6-Chapter Manuscript Tutorial**](/word/) — Follow the full workflow from prose figures to ACS B&W submission.
