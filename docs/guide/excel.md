# SMILES for Microsoft Excel

**SMILES for Microsoft Excel** renders high-resolution 2D chemical structure graphics directly inside worksheet cells with instant in-cell formula recalculation and seamless double-click editing.

---

## 🎥 3-Minute Video Walkthrough

> 💡 **Master all core features in under 2 minutes (1080p with English & Korean subtitles).**  
> Covers in-cell formula entry, chemical styling options, offline Ketcher 2D editor workflows, reaction formulas (`=KET2IMG`), and one-click property calculation.

<p align="center">
  <video controls="controls" width="100%" style="max-width: 800px; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.15);" poster="/excel-showcase.png">
    <source src="/videos/excel-guide-video.mp4" type="video/mp4" />
    <source src="https://github.com/user-attachments/assets/e566a830-55d5-4ff5-8030-344d8b8ce7b6" type="video/mp4" />
    <a href="/videos/excel-guide-video.mp4">🎬 Watch Excel 3-Minute Video Walkthrough</a>
  </video>
</p>

<details>
<summary><b>⏱️ Video Chapters & Timeline</b></summary>

| Time | Chapter | Highlights |
| :---: | :--- | :--- |
| **00:00** | **Overview & Architecture** | Native in-cell images, 100% offline local rendering, auto-scaling with row/column size |
| **00:13** | **Step 1: In-Cell Formulas** | Enter `=SMILES2IMG(A2)` and fill down hundreds of compounds instantly |
| **00:28** | **Step 2: Chemical Styling** | Monochrome patent line art (`bw`), NMR peak indices (`num`), stereochemistry (`stereo`), dark mode & rotations |
| **00:44** | **Step 3: Ketcher 2D Editor** | Shortcut `Ctrl+Shift+K` or right-click to open Ketcher and apply edits directly back to cell |
| **01:00** | **Step 4: Chemical Reactions** | Multi-component reactions with arrows and annotations using `=KET2IMG(A2)` |
| **01:15** | **Step 5: Quick Insert & Properties** | Clipboard paste and instant Chemical Formula / MW calculation via context menus |
| **01:30** | **Summary & Checklist** | Best practices for research spreadsheets and publication tables |

</details>

---

## 🌟 Core Highlights

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" />
</p>

::: tip Quick Summary
1. **Draw in an Empty Cell:** Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>, draw structure, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
2. **Formula Mass Insertion:** Type `=SMILES2IMG(A2)` or `=KET2IMG(A2)` and drag down.
3. **Quick Insert from Clipboard:** Copy SMILES/KET and click **[SMILES] ➔ [Insert from Clipboard]**.
4. **Edit Anytime:** Double-click the cell to open the Ketcher editor, or right-click to copy molecular properties.
:::

---

## 📝 Functions: Native In-Cell vs. Floating Images

Depending on your Microsoft Excel version, two distinct formula modes are available:

### 1. `=SMILES2IMG(smiles_or_ket, [bg], [options])`
* **Target:** **Microsoft 365** and **Excel 2024+**
* Embeds real **native in-cell pictures** directly into cell contents.
* Moves, sorts, filters, and auto-scales with row height and column width.
* Can be referenced by standard Excel formulas and formatting rules.

### 2. `=SMILES2IMG.FLOAT(smiles_or_ket, [bg], [options])`
* **Target:** **Excel 2016 / 2019 / 2021** (or when floating shapes are explicitly desired)
* Inserts floating picture shapes locked to cell coordinates (`xlMoveAndSize`).
* Automatically recalculates, replaces existing shapes without orphan clutter, and moves/resizes with cell boundaries.

---

## 🎨 Chemical Styling & Rendering Options

Options can be combined freely using spaces, commas, semicolons, or pipe symbols:

| Option Keyword | Description | Example |
| :--- | :--- | :--- |
| `bw` / `monochrome` | High-contrast black-and-white line art (ACS / USPTO patent standard) | `=SMILES2IMG(A2, "", "bw")` |
| `colored` | Standard IUPAC element color-coded atoms | `=SMILES2IMG(A2, "", "colored")` |
| `num` / `numbers` | Displays atom index numbers (ideal for NMR/MS peak assignments) | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | Forces CIP stereochemistry labels (*R/S*, *E/Z*) | `=SMILES2IMG(A2, "", "stereo")` |
| `d=2.0` / `d=1.5` | High-DPI supersampling multiplier for ultra-sharp rasterization | `=SMILES2IMG(A2, "", "d=2.0")` |
| `flip` / `flipv` | Flips structure horizontally or vertically | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` | Rotates molecule in 90-degree increments | `=SMILES2IMG(A2, "", "rot90")` |
| `w=300 h=200` | Custom pixel dimensions for floating images | `=SMILES2IMG.FLOAT(A2, "", "w=300 h=200")` |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" />
</p>

---

## 🧪 Lossless Reaction Canvases: `=KET2IMG`

For multi-component organic syntheses involving reaction arrows, catalysts, yields, and plus signs:

```text
=KET2IMG(A2)
```
* **Lossless Preservation:** The full reaction graph and coordinate layout in KET format are preserved in the cell formula.
* **Double-Click In-Place Re-editing:** Double-clicking the cell opens Ketcher with all reaction arrows, plus signs, and notes fully editable.

---

## 📋 Context Menu & Molecular Property Extraction

Right-click any cell containing chemical structure formulas or text:

* **Chemical Structure ➔ Edit Structure in Ketcher:** Launches the offline editor.
* **Chemical Structure ➔ Copy SMILES / Copy KET:** Extracts exact structure strings to the clipboard.
* **Chemical Structure ➔ Copy Molecular Formula:** Calculates Hill-system chemical formula (e.g. `C8H10N4O2`).
* **Chemical Structure ➔ Copy Molecular Weight (MW):** Calculates molecular weight in g/mol (e.g. `194.19`).
* **Chemical Structure ➔ Copy Exact Mass:** High-resolution monoisotopic mass (e.g. `194.0804`).
