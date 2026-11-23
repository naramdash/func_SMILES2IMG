# 1. 2D Structures & Reaction Schemes

Sketch 2D chemical structures and multi-step reaction schemes directly in Microsoft Word with offline **Ketcher 2D**, convert typed SMILES, InChI, or compound names with **`Paste Structure`**, and double-click any structure anytime to re-edit without shifting paragraph layout.

## 1. Choose an Input Method

| Starting From | Recommended Action | Result in Word |
| :--- | :--- | :--- |
| **Blank cursor** (drawing a new molecule or reaction scheme) | Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (or right-click ➔ **`Chemical Structure` ➔ `Insert / Edit Structure...`**) | Opens offline **Ketcher 2D**; press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert an aspect-ratio-locked PNG |
| **Typed text in Word** (`CC(=O)Oc1ccccc1C(=O)O`, `InChI=...`, `Aspirin`, `50-78-2`) | Highlight the text ➔ click **`[SMILES]` ➔ `[Paste Structure]`** | Replaces the highlighted text **in place** with the rendered structure |
| **Copied string on Clipboard** (from PubChem, Excel, ChemDraw, or ELN) | Place cursor without highlighting text ➔ click **`[Paste Structure]`** | Renders the clipboard SMILES, InChI, KET, or compound name at the cursor |
| **Typed text you want to modify before inserting** | Highlight `Aspirin` or a SMILES string ➔ click **`[Insert / Edit Structure]`** | **Preloads** that exact molecule straight into Ketcher 2D so you can modify substituents immediately |

::: tip Power Tip: Preload Text Directly Into Ketcher
You do not have to insert first and double-click later. Whenever a chemical name, valid CAS number, or SMILES string is already typed in your draft, highlight it and click **`[Insert / Edit Structure]`**—Ketcher opens with that molecule pre-drawn on the canvas. *(Ambiguous short English words are safely skipped during preloading so ordinary prose never preloads unintended molecules.)*
:::

## 2. Sketching Inline Structures & Text Conversion

### Sketch in Place (`Ketcher 2D`)

<FeatureMedia id="word-ch1-insert-ketcher" />

1. Place your cursor inside a Word paragraph or table cell.
2. Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** and draw your molecule in **Ketcher 2D** (see [Ketcher 2D Manual](/editor/) for full drawing shortcuts and background toggles).
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**) to insert a crisp, high-DPI inline picture (`InlineShape`) with `LockAspectRatio` locked automatically.

### Convert Typed SMILES, InChI, or Chemical Names (`Paste Structure`)

<FeatureMedia id="word-ch2-quick-insert" />

* **Supports SMILES, InChI, KET, Names & CAS:** Common drugs, solvents, and bioactives (`Aspirin`, `Caffeine`, `Paracetamol`, `THF`, `50-78-2`) resolve offline from the built-in dictionary; unknown names and valid CAS numbers resolve via PubChem and cache locally on disk for **30 days** (`%LOCALAPPDATA%\Smiles2Img`). Invalid CAS check digits are rejected locally, and compound names or ambiguous short English words show a confirmation dialog on **`[Paste Structure]`**.
* **Replaces Highlighted Text In-Place:** Highlighting `Aspirin` or a SMILES string in prose or a table cell and clicking **`[Paste Structure]`** swaps the text directly for the rendered structure.

### Inline vs. Floating Pictures in Word

* **Inline (`InlineShape`, default):** Anchored inside the text line—best for prose mentions and compound table cells because it never drifts across pages when you edit preceding paragraphs.
* **Floating (`Shape`, after applying Text Wrapping):** If you switch a structure to *Square* or *In Front of Text* for multi-column callouts, `func_SMILES2IMG` intercepts Word's double-click so it still reopens Ketcher 2D in place instead of Word's Format Picture pane.

## 3. Multi-Step Synthetic Reaction Schemes (`KET`)

Draw complete synthesis schemes with reaction arrows (`➔`), `+` signs, reagents, catalysts, temperatures, and yields—and preserve every coordinate losslessly inside `.docx`:

<FeatureMedia id="word-ch4-reaction-scheme" />

1. Place your cursor on a dedicated scheme paragraph or wide figure cell and click **`[SMILES]` ➔ `[Insert / Edit Structure]`**.
2. In **Ketcher 2D**, draw reactants, products, reaction arrows (`➔`), `+` signs, and text annotations (`T` tool for conditions such as `H₂SO₄, 85 °C, 92%`).
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**).
4. **Lossless `.docx` Persistence:** Whenever arrows, `+` signs, or multi-component layouts are present, `func_SMILES2IMG` saves the full KET JSON coordinates inside the picture's **Alt Text** (`KET:...`) and canonical SMILES in `Title`. Even after saving, closing, or sharing the `.docx`, **double-clicking** the scheme on any PC with `func_SMILES2IMG` restores every arrow angle and catalyst label intact.

## 4. Double-Click Re-Editing & Cross-App Migration

* **Layout-Safe Replacement:** Double-clicking any inline or floating 2D picture reopens Ketcher 2D. Pressing <kbd>Ctrl</kbd> + <kbd>Enter</kbd> swaps the updated image at the exact original width and wrapping mode without disrupting surrounding text.
* **From Excel to Word:** Select a formula cell in Excel (`=SMILES2IMG(...)` or `=KET2IMG(...)`), click **`[Copy KET]`** (for reactions) or **`[Copy SMILES]`**, then place your cursor in Word and click **`[Paste Structure]`**.
* **From ChemDraw to Word:** Select your structure or reaction in ChemDraw, choose **Edit ➔ Copy As ➔ SMILES** (or paste into Ketcher / use **`[Paste Structure]`** in Word) to turn legacy figures into double-click-editable Word objects.

## Related Guides

- [**2. SAR Tables & Experimental Properties**](/word/tables-properties) — Prevent clipped inline pictures in tables and copy Formula, MW, and HRMS Exact Mass accurately
- [**Ketcher 2D Sketcher Manual**](/editor/) — Full drawing tools, ring templates, and keyboard shortcuts
- [**Double-Click & Editing Troubleshooting**](/word/troubleshooting/double-click) — What to check if double-clicking opens Word's Picture Format pane
