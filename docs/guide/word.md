# SMILES for Microsoft Word

**SMILES for Microsoft Word** brings seamless chemical structure editing and inline high-resolution molecule rendering directly into Word documents, research manuscripts, journal articles, and thesis templates.

---

## 🌟 Visual Feature Walkthrough

All chemical workflows in Microsoft Word are designed to behave naturally like text characters—wrapping inline within sentences, aligning inside table cells, and editing in-place without disturbing document layout:

<p align="center">
  <img src="/word-showcase.png" alt="SMILES for Word Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip Key Advantages in Word
* **Inline Picture Flow:** Structures are inserted as native `InlineShapes` that behave like characters. They wrap naturally with sentences, align within table cells, and move with paragraphs.
* **Layout Preservation:** When double-clicking to edit a structure, the newly generated image replaces the old shape in-place without distorting table column widths or paragraph margins.
* **Lossless File Portability:** All structure metadata (`KET` or `SMILES`) is stored inside the `.docx` XML structure (`AlternativeText` and `Title`). Sharing the file with colleagues keeps all chemical metadata intact.
:::

---

## 📑 Dedicated Office Ribbon: `[SMILES]` Tab

The `[SMILES]` tab on Word's ribbon provides one-click access to structure drawing, instant text replacement, and diagnostics:

<p align="center">
  <img src="/ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ Insert Structure (`Ctrl+Shift+K`):** Launches the offline Ketcher 2D editor. Draw your molecule or multi-step reaction, then press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**) to insert it as an inline picture at the current cursor position.
2. **⚡ Quick Insert:**
   - If text containing a SMILES string is highlighted in your document, converts it directly into an inline chemical structure in 0.1 seconds.
   - If no text is selected, parses the clipboard and inserts the structure at the cursor position.
3. **📋 Insert from Clipboard:** Reads clipboard content as SMILES or KET and embeds it as an inline shape.
4. **🔍 Settings & Diagnostics:** Displays COM add-in status, chemical engine version, and troubleshooting logs.

---

## 🖱️ Redesigned Context Menu: `Chemical Structure`

Right-click any chemical structure picture or paragraph in your Word document:

<p align="center">
  <img src="/context-menu-guide.png" alt="Chemical Structure Context Menu in Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Case A: Right-Clicking an Existing Chemical Picture
* **✏️ Edit Structure in Ketcher:** Opens the offline editor with the exact molecule or reaction graph loaded.
* **📋 Copy SMILES / Copy Molfile (KET):** Copies clean chemical notations to the clipboard.
* **🔬 Copy Molecular Formula:** Calculates the standard Hill-notation formula (e.g. `C8H10N4O2`).
* **⚖️ Copy Molecular Weight (MW):** Copies IUPAC molecular weight in g/mol (e.g. `194.19`).
* **🎯 Copy Exact Mass:** Copies high-resolution monoisotopic mass (e.g. `194.0804`).

### Case B: Right-Clicking Document Text or Blank Paragraph
* **✏️ Insert Structure in Ketcher:** Launches Ketcher to draw and insert a new structure at the cursor.
* **📋 Paste from Clipboard:** Reads clipboard chemical text and embeds it as an inline image.

---

## 🧪 Ketcher 2D Molecule & Reaction Editor

Double-click any chemical structure picture inside Word or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> to launch Ketcher:

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **Real-Time Property Calculation:** The top status bar calculates Formula, MW, and Exact Mass in real time as you draw.
* **Transparent Background Toggle:** Switch between white and transparent alpha backgrounds.
* **Table Cell Auto-Fit:** The rendered picture is sized appropriately for document flow and table cells.
* **Keyboard Hotkeys:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to apply updates to Word in-place, or <kbd>Esc</kbd> to cancel.

---

## 🔬 Multi-Step Reactions & Schemes in Manuscripts

For reaction schemes, organic syntheses, and schemes with arrows and catalyst annotations:

* **Full Scheme Preservation:** Reaction arrows, catalyst conditions, and plus signs are stored losslessly inside the document's shape metadata (`AlternativeText`).
* **Double-Click Re-editing:** Double-clicking the reaction scheme in your Word document reopens Ketcher with the full reaction graph intact—enabling painless updates before journal submission.

---

## 🔍 Built-In Diagnostics Console

Click **[SMILES] ➔ [Settings & Diagnostics]** to inspect engine health:

* Confirms EPAM Indigo C++ native library initialization and COM add-in registration.
* Validates WebView2 runtime and offline Ketcher bundle assets.
* Displays real-time rendering logs and timing benchmarks.
