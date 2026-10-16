---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "Free & Open-Source Office Chemical Add-in Suite"
  tagline: "Publication-quality 2D chemical rendering (SMILES · InChI · KET · Name/CAS), embedded Ketcher 2D sketcher, Mol* 3D macromolecular viewer, SAR & drug-likeness formulas, Vector SVG copy, Batch Style normalizer, and SDF I/O across Excel, Word, and PowerPoint — 100% local & offline."
  actions:
    - theme: brand
      text: Get Started
      link: /guide/getting-started
    - theme: alt
      text: Excel Guide
      link: /guide/excel
    - theme: alt
      text: Formula Reference
      link: /reference/formula
    - theme: alt
      text: View on GitHub
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: Real-Time In-Cell Formulas & Highlighting
    details: Render structures directly into Excel cells via =SMILES2IMG, =INCHI2IMG, =KET2IMG, or =NAME2IMG("Aspirin"). Supports SMARTS substructure highlighting ("hl:c1ccccc1") and auto-scales with cells.
  - icon: 🔬
    title: Cheminformatics SAR & Drug-Likeness
    details: Screen compound libraries in 0ms purely offline with =SMILES.MATCH, =SMILES.SIMILARITY, =SMILES2LIPINSKI, =SMILES2LOGP, =SMILES2TPSA, =SMILES2EXACTMASS, =SMILES2MW, and =SMILES2FORMULA.
  - icon: ✏️
    title: Instant Ketcher 2D Sketcher (<0.05s)
    details: Double-click any molecule or reaction in Excel, Word, or PowerPoint to edit in-place with EPAM Ketcher v3.18.0, powered by warm WebView2 instance pooling for instant reopening.
  - icon: 🧬
    title: Mol* 3D Macromolecular Viewer
    details: Visualize proteins, nucleic acids, and complexes across Excel, Word, and PowerPoint. Download PDB IDs directly and insert rotatable 3D GLB models or high-res 2D pictures.
  - icon: 🖼️
    title: Vector SVG Copy & Batch Style Normalizer
    details: Copy crisp scalable vector graphics (SVG) directly to the clipboard for Illustrator/Inkscape, or convert all structures in a sheet, document, or slide deck to ACS B&W or Dark Mode in one click.
  - icon: 📂
    title: Native Excel SDF Import & Export
    details: Import multi-molecule .sdf chemical files directly into formatted Excel tables with live structure and property formulas, or export worksheet libraries back to .sdf.
  - icon: 🛡️
    title: 100% Local & Privacy-Safe
    details: Powered by embedded EPAM Indigo, Ketcher, and Mol* runtimes. Zero cloud telemetry, zero data leakage, and lossless CustomXMLParts document persistence.
  - icon: 📦
    title: One-Line Installer & Auto Update Check
    details: A single PowerShell command installs, unblocks, and registers the entire suite for Excel, Word, and PowerPoint without admin rights, complete with an in-app update checker.
---

## ⚡ Quick Install

Install the entire suite (Excel, Word, PowerPoint) in seconds via **PowerShell**:

```powershell
# Stable release (v1.6.0)
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex

# Preview release (v1.7.0-preview.1 — SAR, Drug-Likeness, Name2Img, SVG Copy, Batch Style & SDF I/O)
$v='v1.7.0-preview.1'; irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 Automatic Upgrade & In-App Update Check
Running the install command anytime upgrades your installation to the latest release, or click **[SMILES] ➔ [About SMILES] ➔ [Check for Updates]** directly inside Excel, Word, or PowerPoint!
:::

---

## 🖼️ Application Showcase

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel](/guide/excel) | [📘 Microsoft Word](/guide/word) | [📙 Microsoft PowerPoint](/guide/powerpoint) |
| :---: | :---: | :---: |
| In-cell formulas (`=SMILES2IMG`, `=NAME2IMG`), SAR & Lipinski screening, `.sdf` table import/export, Batch Style & SVG copy | Inline shapes flowing with paragraphs, table layout preservation, Mol* 3D models, Batch Style & Vector SVG copy | One-click slide placement, dark theme transparent backgrounds, rotatable 3D GLB models, presentation-wide Batch Style |

| [✏️ Ketcher 2D Editor](/guide/ketcher) | [🧬 Mol* 3D Viewer](/guide/molstar) | [📐 Formula Reference](/reference/formula) |
| :---: | :---: | :---: |
| `<0.05s` warm-pooled 2D molecule & reaction sketcher, lossless KET format | Macromolecular 3D viewer across Excel, Word & PowerPoint, PDB download & GLB 3D | Complete reference for 25+ rendering, property, Lipinski & SAR worksheet functions |
