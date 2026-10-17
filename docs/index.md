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
      text: 🔬 SAR & Lipinski Guide
      link: /guide/excel-cheminformatics
    - theme: alt
      text: 🛠️ Formula Generator
      link: /tools/formula-builder

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

---

## 🖥️ Explore by Office Application (Click a Tab!)

<OfficeAppTabs />

---

## 🎯 Jump Straight to What You Need

| What You Want to Do | Fastest Way | Detailed Guide |
| :--- | :--- | :---: |
| **Install & render your very first molecule** | Run the 1-line PowerShell installer, then try `=NAME2IMG("Aspirin")` or <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd> | [1-Min Quickstart ➔](/guide/getting-started) |
| **Build a structure table in Excel from names or SMILES** | Type `=SMILES2IMG(A2)` or `=NAME2IMG("50-78-2")` and drag down | [Excel Basics ➔](/guide/excel) |
| **Calculate MW, Exact Mass, Lipinski Ro5, or screen SAR** | Use `=SMILES2MW`, `=SMILES2LIPINSKI`, `=SMILES.MATCH`, and `.sdf` import/export | [Excel SAR & SDF ➔](/guide/excel-cheminformatics) |
| **Insert structures into Word papers or PowerPoint slides** | Inline text wrapping (Word), transparent dark mode & live 3D rotation (PPT) | [Word ➔](/guide/word) · [PowerPoint ➔](/guide/powerpoint) |
| **Convert to ACS B&W, copy Illustrator SVG, or export tables** | Use `[Batch Style]`, `[Copy SVG]`, and `[Molecule Info]` multi-format exports | [Publishing & SVG ➔](/guide/publishing) |
| **Sketch 2D reaction schemes or load 3D proteins** | Built-in Ketcher 2D Sketcher (`=KET2IMG`) and Mol* 3D Viewer (`.glb`) | [2D Sketcher ➔](/guide/ketcher) · [3D Viewer ➔](/guide/molstar) |
| **Look up all 25+ Excel functions and argument keywords** | Browse task-organized reference tables and style aliases | [Formula Cheat Sheet ➔](/reference/formula) |
| **Build & preview custom Excel formulas interactively** | Click options in the web generator widget to preview & copy in 1 click | [🛠️ Formula Generator ➔](/tools/formula-builder) |
