---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "Chemical Structures & 3D Proteins inside Microsoft Office"
  tagline: "Draw 2D molecules and reaction schemes, calculate molecular weight & drug-likeness via Excel formulas, and embed rotatable 3D proteins directly inside Excel, Word, and PowerPoint — 100% free and offline."
  actions:
    - theme: brand
      text: 🚀 Install & 1-Min Quickstart
      link: /guide/getting-started
    - theme: alt
      text: 📗 Excel Guide
      link: /guide/excel
    - theme: alt
      text: 📐 Formula Cheat Sheet
      link: /reference/formula
    - theme: alt
      text: GitHub
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: "Auto-Render Structures in Excel Cells"
    details: "Turn chemical names, CAS numbers, or SMILES into crisp in-cell structure images with =SMILES2IMG(A2) or =NAME2IMG('Aspirin')—and drag down across hundreds of rows."
  - icon: ✏️
    title: "Draw & Edit Anytime by Double-Clicking"
    details: "Double-click any cell or picture (or press Ctrl+Shift+K) to open the built-in Ketcher 2D chemical sketcher for molecules and multi-step reaction schemes."
  - icon: 🔬
    title: "Instant MW, Formula & Drug-Likeness Formulas"
    details: "Calculate molecular weight, Hill formula, LogP, TPSA, Lipinski Rule of 5, and substructure matches right in your sheet with =SMILES2MW, =SMILES2LIPINSKI, and =SMILES.MATCH."
  - icon: 🧬
    title: "Interactive 3D Proteins (Mol*)"
    details: "Type a 4-letter PDB ID (e.g., 1CRN, 4HHB) or open a local file to insert 360°-rotatable 3D protein models or high-res 2D figures into Excel, Word, and PowerPoint."
  - icon: 🎨
    title: "1-Click ACS B&W, Dark Mode & Vector SVG"
    details: "Convert all structures in a sheet, document, or slide deck to ACS monochrome or dark-mode white ink in one click, copy lossless SVG vectors for Illustrator, or import/export .sdf files."
  - icon: 🛡️
    title: "100% Offline & Confidential"
    details: "All chemical rendering and property calculations happen locally on your PC. Proprietary drug candidates and unpublished structures never leave your machine."
---

## ⚡ 10-Second Quick Install (Excel · Word · PowerPoint)

Open **PowerShell** and paste this single command (no Administrator rights needed):

```powershell
# Stable release (v1.7.0) — installs or upgrades automatically
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 Try It Immediately After Installing!
Restart Excel, Word, or PowerPoint and look for the **`[SMILES]`** tab on the top ribbon—or type `=NAME2IMG("Aspirin")` into any Excel cell! Follow the **[1-Minute Quickstart](/guide/getting-started)** for a step-by-step tour.
:::

---

## 🎯 What Would You Like to Do?

Jump straight to the workflow you need:

| Goal | Fastest Way | Guide |
| :--- | :--- | :---: |
| **Build a chemical table in Excel from names or SMILES** | Type `=NAME2IMG("Caffeine")` or `=SMILES2IMG(A2)` and drag down | [Excel Guide ➔](/guide/excel) |
| **Draw a molecule or reaction scheme by hand** | Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> in Excel, Word, or PowerPoint | [2D Sketcher ➔](/guide/ketcher) |
| **Calculate MW, Formula, LogP, or Lipinski Ro5** | Use `=SMILES2MW(A2)`, `=SMILES2FORMULA(A2)`, `=SMILES2LIPINSKI(A2)` | [Formula Cheat Sheet ➔](/reference/formula) |
| **Write a paper in Word or make slides in PowerPoint** | Insert inline structures, apply 1-click ACS B&W or Dark Mode themes | [Word ➔](/guide/word) · [PowerPoint ➔](/guide/powerpoint) |
| **Insert a rotatable 3D protein structure** | Click **`[SMILES] ➔ [3D Structure (Mol*)]`** and enter a PDB ID (e.g., `1CRN`) | [3D Viewer ➔](/guide/molstar) |

---

## 🖼️ Office Application Showcase

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel](/guide/excel) | [📘 Microsoft Word](/guide/word) | [📙 Microsoft PowerPoint](/guide/powerpoint) |
| :---: | :---: | :---: |
| In-cell structure formulas, MW/Lipinski/SAR screening, substructure highlighting, `.sdf` file import & export | Inline structures that flow with sentences and tables, 1-click ACS B&W manuscript styling, instant property table paste | 1-click slide placement, transparent backgrounds for dark themes, 360°-rotatable 3D protein models |
