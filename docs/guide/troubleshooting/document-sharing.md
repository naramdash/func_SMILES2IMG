# Document Sharing

How `.docx`, `.pptx`, and `.xlsx` files created with **`func_SMILES2IMG`** behave when opened on **macOS**, iPad, Office for the Web, or PCs without the add-in installed.

## 1. Microsoft Word (`.docx`) & PowerPoint (`.pptx`)

* **Viewable Across Platforms:** 2D chemical structures are embedded as standard high-DPI PNG images, and 3D proteins are embedded as standard `.glb` 3D models.
* **No Add-In Needed for Readers:** Colleagues on Windows, **macOS**, iPad, or Office for the Web can view, print, present, and export to PDF without missing images or broken layouts.
* **Round-Trip Editability Preserved:** Because chemical metadata (`KET:` / `SMILES:` / `MOLSTAR:`) travels inside standard Office Alt Text and `CustomXMLParts`, reopening the same file later on a PC with `func_SMILES2IMG` restores full double-click editing automatically.

![Embedded 2D structures and 3D figures in Microsoft Word](/word-showcase.png)

## 2. Microsoft Excel (`.xlsx`)

* **Cached Outputs Stay Visible:** `.FLOAT` picture shapes, Microsoft 365 in-cell image caches, and calculated descriptor values (`MW`, `Formula`, `LogP`, `TPSA`) remain visible on other PCs as long as the workbook is not recalculated.
* **Preventing `#NAME?` on External Sharing:** If a recipient without `func_SMILES2IMG` double-clicks a custom formula cell or triggers a full recalculation (<kbd>F9</kbd>), Excel displays `#NAME?` for custom functions.
* **Recommended Final-Export Workflow:**
  1. Keep your original working `.xlsx` with live formulas.
  2. For external recipients without the add-in, either copy numeric/text formula columns (<kbd>Ctrl</kbd> + <kbd>C</kbd>) ➔ **Paste as Values**, or export the sheet to **PDF**.

## Related Troubleshooting

- [**Word Double-Click Editing**](/word/troubleshooting/double-click) — How Alt Text metadata preserves editability across shared Word & PPT files.
- [**Frequently Asked Questions (FAQ)**](/guide/faq) — ACS B&W styling, lossless SVG export, .sdf database files, and updates.
