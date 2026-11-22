# Table Clipping & Batch Style Scope

Solutions when inline chemical structures look sliced off at the top edge inside Word paragraphs or table cells, when column widths shift during edits, or when `Batch Style` updates a different scope than expected.

## 1. Top of Inline Structure Is Cut Off in a Line or Table Cell {#table-layout}

When an inline picture (`InlineShape`) appears cleanly at the bottom half while its top bonds disappear under the previous line:

* **Root Cause:** The host paragraph or table style has **Line Spacing** locked to **`Exactly`** (for example, *Exactly 12 pt*), capping rendered line height strictly at the font box height.
* **Fix:** Select the affected structure paragraph or table column ➔ **Home ➔ Paragraph ➔ Line Spacing Options** ➔ change Line Spacing to **Single** or **Multiple** (`1.15`). The full 2D structure expands visibly right away.

## 2. Table Column Widths Shift After Re-Editing

* **Lock table widths before inserting structures:** Right-click the table ➔ **AutoFit ➔ Fixed Column Width**. `func_SMILES2IMG` automatically keeps aspect ratio locked (`LockAspectRatio`) and replaces structures inside your existing cell bounds.
* **Low contrast on shaded rows or alternating stripes:** Run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Standard Color (Transparent)]`** so structures sit cleanly over cell shading without white rectangles, or pick **`[White Background (Color)]`** to place a crisp white card behind each molecule.

## 3. Batch Style Changed Too Few or Too Many Figures {#batch-style}

`Batch Style` scopes itself automatically based on what is currently selected in Word:

* **Pictures, text, or table cells highlighted:** Converts **only the structures inside that selection**.
* **Nothing selected (text cursor blinking in prose):** Converts **every 2D molecule and `KET` reaction scheme across the entire `.docx` document**, leaving `Mol*` 3D protein figures untouched so colored protein ribbons remain intact.

::: tip Check Selection Before Clicking Batch Style
If you intended to restyle your whole manuscript and only one structure changed, click empty prose with nothing highlighted and run **`[Batch Style]`** again.
:::

## Related Guides

- [**Double-Click & Editing**](/word/troubleshooting/double-click) — Restore double-click editing when Word Picture Format opens instead of Ketcher or Mol\*
- [**2. SAR Tables & Experimental Properties**](/word/tables-properties) — Layout-safe fixed-width SAR tables and 17-descriptor `Molecule Info` copy
- [**4. Journal Finish, SVG, Sharing**](/word/publishing) — Pre-submission ACS monochrome conversion and Illustrator vector `SVG` export
