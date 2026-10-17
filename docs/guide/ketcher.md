# Ketcher 2D Sketcher Guide

Learn how to sketch molecules and multi-step organic reaction schemes directly inside Excel, Word, and PowerPoint using the built-in **Ketcher 2D Editor**.

<p align="center">
  <img src="/ketcher-guide.png" alt="Ketcher Editor Interface" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip 💡 Why You'll Love It
* **100% Offline & Private:** Runs entirely on your local PC without internet access—keeping unpublished structures completely confidential.
* **Instant Reopen (<0.05s):** Opens instantaneously whenever you double-click a structure to make quick edits.
* **Live MW & Formula Calculation:** Automatically displays **Formula, Molecular Weight (MW), and Exact Mass** in the top bar while you draw.
:::

---

## 1. Draw & Edit in 3 Steps

<p align="center">
  <img src="/workflow-ketcher.png" alt="Ketcher Workflow" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Step 1. Open the Sketcher
* **Draw a New Structure:** Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>, or click **`[SMILES]`** ➔ **`[Insert / Edit (SMILES)]`** in Excel (or **`[Insert / Edit Structure]`** in Word & PowerPoint) on the top ribbon.
* **Edit an Existing Structure:** Simply **double-click** any chemical cell in Excel or structure picture in Word / PowerPoint.

### Step 2. Sketch Molecules or Reaction Schemes
* **Bonds & Ring Templates (Left Toolbar):** Place single, double, triple, or stereo bonds (wedges/hashes) and common ring templates (benzene, cyclohexane, cyclopentane).
* **Elements (Right Toolbar):** Pick common atoms (`C`, `N`, `O`, `S`, `P`, halogens) or open the full periodic table.
* **Multi-Step Reactions:** Use the reaction arrow (`➔`), plus sign (`+`), and text (`T`) tools to add catalysts (H₂SO₄), temperatures (55 °C), and reaction yields.

### Step 3. Choose Background & Apply (`Ctrl + Enter`)
* In the top bar, toggle between **`[Background: White]`** (for standard pages and sheets) and **`[Background: Transparent]`** (for dark slides and colored cells).
* Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**) to insert the rendered structure straight into your document!

---

## 2. How Single Molecules vs. Multi-Step Reactions Are Saved

Ketcher automatically picks the best lossless format based on what you draw:

| Feature | Single Molecule Sketch | Multi-Step Reaction Scheme (with Arrows · `+` · Text) |
| :--- | :--- | :--- |
| **Auto-Detected When** | Only a single chemical structure is drawn | Reaction arrows (`➔`), `+` signs, or catalyst/condition labels are present |
| **Excel Formula** | `=SMILES2IMG("...")` | `=KET2IMG("...")` (preserves 2D layout, arrows, and text labels 100%) |
| **Word / PPT Picture** | High-DPI image + embedded SMILES metadata | High-DPI image + embedded KET JSON metadata |
| **Double-Click Re-Edit** | Restores the exact molecule | Restores every arrow angle, reagent label, and coordinate |

---

## 3. Essential Shortcuts

| Shortcut / Action | Function | Description |
| :--- | :--- | :--- |
| <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> | **Open Sketcher** | Open Ketcher for the selected cell or cursor position |
| **Double-Click** | **Edit In-Place** | Reopen any existing molecule or reaction scheme to edit |
| <kbd>Ctrl</kbd> + <kbd>Enter</kbd> | **Apply to Document** | Insert the drawing into Excel/Word/PowerPoint and close |
| <kbd>Ctrl</kbd> + <kbd>Z</kbd> / <kbd>Y</kbd> | **Undo / Redo** | Undo or redo recent bond/atom edits |
| <kbd>Esc</kbd> | **Cancel** | Close the editor without saving changes |

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../guide/molstar" class="s2i-card">
    <div class="s2i-card-title">🧬 Mol* 3D Protein & Crystal Viewer</div>
    <p class="s2i-card-desc">Load PDB structures (e.g., 1CRN, 6LU7) and capture 3D protein ribbons & binding pockets</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 Publishing, SVG & Property Tables</div>
    <p class="s2i-card-desc">Convert drawn molecules to ACS B&W or copy lossless SVG vectors for Illustrator</p>
  </a>
</div>

