---
prev:
  text: Getting Started
  link: /guide/getting-started
next:
  text: 3D Protein Viewer & GLB Models
  link: /editor/molstar
---

# Draw 2D Structures & Reaction Schemes (Ketcher)

Draw molecules and multi-step reaction schemes from scratch from the **`[SMILES]`** ribbon tab, or **double-click any existing structure** in Excel, Word, or PowerPoint to reopen and edit it in place in 0.05 seconds.

<FeatureMedia id="common-ketcher-2d" />

::: tip 💡 Two-Way Round-Trip Editing Across Excel, Word & PowerPoint
* **Double-Click Re-Editing (<0.05s):** Every structure inserted into an Excel cell, Word page, or PowerPoint slide retains its lossless chemical coordinates—just **double-click** it anytime to modify atoms, bonds, or reaction arrows!
* **Top Toolbar & Live Property Lookup:** Use **`Background: White` / `Background: Transparent`**, **`Cancel (Esc)`**, and **`Apply (Ctrl+Enter)`** in the top toolbar; inspect live Formula, MW, and Exact Mass anytime via Ketcher's built-in **Calculated Values** tool or Office's **`[Molecule Info]`** menu.
* **100% Offline & Confidential:** Runs entirely on your local PC without internet access.
:::

## 1. Draw & Edit in 3 Steps

### Step 1. Open the Sketcher
* **Draw a New Structure:** Click **`[SMILES]`** ➔ **`[Insert / Edit (SMILES)]`** in Excel (or **`[Insert / Edit Structure]`** in Word & PowerPoint) on the top ribbon.
* **Edit an Existing Structure:** Simply **double-click** any chemical cell in Excel or structure picture in Word / PowerPoint.

### Step 2. Sketch Molecules or Reaction Schemes
* **Bonds & Ring Templates (Left Toolbar):** Place single, double, triple, or stereo bonds (wedges/hashes) and common ring templates (benzene, cyclohexane, cyclopentane).
* **Elements (Right Toolbar):** Pick common atoms (`C`, `N`, `O`, `S`, `P`, halogens) or open the full periodic table.
* **Multi-Step Reactions:** Use the reaction arrow (`➔`), plus sign (`+`), and text (`T`) tools to add catalysts (H₂SO₄), temperatures (55 °C), and reaction yields.

### Step 3. Choose Background & Apply (`Ctrl + Enter`)
* In the top bar, toggle between **`[Background: White]`** (for standard pages and sheets) and **`[Background: Transparent]`** (for dark slides and colored cells).
* Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **`Apply to Cell / Document / Slide`**) to insert the rendered structure straight into your document!

## 2. How Single Molecules vs. Multi-Step Reactions Are Saved

<FeatureMedia id="excel-reactions-ket" />

Ketcher automatically picks the best lossless format based on what you draw:

| Feature | Single Molecule Sketch | Multi-Step Reaction Scheme (with Arrows · `+` · Text) |
| :--- | :--- | :--- |
| **Auto-Detected When** | Only a single chemical structure is drawn | Reaction arrows (`➔`), `+` signs, or catalyst/condition labels are present |
| **Excel Formula** | `=SMILES2IMG("...")` | `=KET2IMG("...")` (preserves 2D layout, arrows, and text labels 100%) |
| **Word / PPT Picture** | High-DPI image + embedded SMILES metadata | High-DPI image + embedded KET JSON metadata |
| **Double-Click Re-Edit** | Restores the exact molecule | Restores every arrow angle, reagent label, and coordinate |

## 3. Essential Actions & Editor Shortcuts

| Action / Shortcut | Function | Description |
| :--- | :--- | :--- |
| **Ribbon `[Insert / Edit]`** | **Open Sketcher** | Open Ketcher for the selected cell or cursor position |
| **Double-Click** | **Edit In-Place** | Reopen any existing molecule or reaction scheme to edit |
| <kbd>Ctrl</kbd> + <kbd>Enter</kbd> | **Apply to Document** | Insert the drawing into Excel/Word/PowerPoint and close |
| <kbd>Ctrl</kbd> + <kbd>Z</kbd> / <kbd>Y</kbd> | **Undo / Redo** | Undo or redo recent bond/atom edits |
| <kbd>Esc</kbd> | **Cancel** | Close the editor without saving changes |

## 🚀 Continue Exploring

- [**Mol\* 3D Protein & Crystal Viewer**](/editor/molstar) — Load PDB structures (e.g., 1CRN, 6LU7) and capture 3D protein ribbons & binding pockets
- [**Publishing, SVG, Property Tables**](/editor/publishing) — Convert drawn molecules to ACS B&W or copy lossless SVG vectors for Illustrator
