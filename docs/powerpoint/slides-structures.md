# 1. 2D Structures & Reaction Slides

Convert PowerPoint TextBoxes containing SMILES strings or compound names in-place in `0.1s`, sketch full multi-step reaction mechanisms in **Ketcher 2D**, and double-click any structure shape anytime to edit in place without losing card alignment.

## 1. Fastest Ways to Put a 2D Molecule on a Slide

| Starting Point on Your Slide | What to Do | How PowerPoint Places It |
| :--- | :--- | :--- |
| **A selected TextBox** containing SMILES, InChI, KET, or name/CAS (`Aspirin`, `Paracetamol`, `50-78-2`) | Click the TextBox border so the shape is selected ➔ **`[SMILES]` ➔ `[Paste Structure]`** | **Replaces the TextBox in-place** at its slide position with an aspect-ratio-locked structure |
| **Highlighted text inside a paragraph** or **string on your Clipboard** | Highlight the text (or copy to clipboard with nothing selected) ➔ **`[Paste Structure]`** | Inserts the rendered structure centered on the slide while keeping surrounding text intact |
| **Typed name/SMILES you want to tweak first** | Select the TextBox or highlight text ➔ **`[Insert / Edit Structure]`** | **Preloads** that molecule into Ketcher 2D immediately so you can adjust groups before inserting |
| **Blank slide** (drawing from scratch) | Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** | Opens offline **Ketcher 2D**; press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert centered on the slide |

::: tip Power Tip: Convert Slide TextBoxes In-Place (`0.1s`)
<FeatureMedia id="powerpoint-slide-textbox-quick-insert" />

When you select a plain-text TextBox (`CC(=O)Oc1ccccc1C(=O)O` or `Paracetamol`) and click **`[Paste Structure]`**, `func_SMILES2IMG` automatically replaces that TextBox right where it sits—expanding thin single-line TextBoxes cleanly into readable bounds while keeping everything inside the slide margins.
:::

## 2. Multi-Step Reaction Mechanism Slides (`KET`)

Build synthetic schemes complete with reaction arrows (`➔`), `+` signs, catalysts, reagents, and yield labels—and keep every arrow and annotation editable inside your `.pptx` file:

<FeatureMedia id="powerpoint-slide-reaction" />

1. Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** on the PowerPoint ribbon (see [Ketcher 2D Manual](/editor/) for drawing shortcuts).
2. Draw reactants, reaction arrows (`➔`), `+` symbols, and text labels (`T` tool for conditions like `H₂SO₄, 55 °C, 92%`).
3. If placing onto a dark or colored slide card, toggle Ketcher's top bar to **`[Background: Transparent]`** before pressing <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Slide`**). *(Or convert all shapes later via `[Batch Style] ➔ [Dark Mode]`.)*
4. **Lossless `.pptx` Persistence:** Full KET JSON coordinates are written into `Shape.Tags["KET"]` and `AlternativeText`, while canonical SMILES goes into `Shape.Tags["SMILES"]` and `Title`. Arrows and labels survive every save, copy-paste across slides, and email share.

## 3. Double-Click In-Place Re-Editing & Cross-App Reuse

* **Keeps Exact Card Position & Size:** Double-clicking any 2D molecule or reaction shape in Normal view reopens Ketcher 2D in `0.05s`. Pressing <kbd>Ctrl</kbd> + <kbd>Enter</kbd> swaps the image right inside the original bounding box (`Left`, `Top`, `Width`, `Height`) so multi-card slide alignment never shifts.
* **Animation Tip:** Because updating a shape replaces the underlying picture object, add PowerPoint **Appear / Fade** animations **after** finalizing your chemical edits (or re-apply them afterward).
* **From Excel or Word to PowerPoint:**
  * **From Word manuscripts:** Select a reaction scheme in Word ➔ **`[Copy KET]`** (or **`[Copy SMILES]`** for a single molecule) ➔ switch to PowerPoint ➔ **`[Paste Structure]`** ➔ **`[Batch Style]` ➔ `[Dark Mode]`**.
  * **From Excel SAR sheets:** Select any `=SMILES2IMG(...)` or `=KET2IMG(...)` cell ➔ **`[Copy SMILES]`** / **`[Copy KET]`** ➔ paste directly onto your slide with **`[Paste Structure]`**.

## Related Guides

- [**2. Dark Mode, Cards, Deck Restyle**](/powerpoint/dark-mode) — Remove white boxes, arrange 4-card SAR grids, and switch the whole deck's theme in 1s
- [**Ketcher 2D Sketcher Manual**](/editor/) — Ring templates, reaction arrows, and keyboard shortcuts
- [**Double-Click & 3D Editing Troubleshooting**](/powerpoint/troubleshooting/3d-editing#double-click) — What to check if double-click opens Format Picture instead of Ketcher
