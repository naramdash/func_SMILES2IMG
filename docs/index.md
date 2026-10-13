---
layout: home

hero:
  name: "SMILES for Office"
  text: "Chemical Structure Renderer & 2D Editor"
  tagline: "High-resolution 2D chemical structure rendering, real-time in-cell formulas, seamless double-click editing, and lossless KET reaction canvases across Excel, Word, and PowerPoint — 100% offline & local."
  image:
    src: /excel-showcase.png
    alt: SMILES for Office
  actions:
    - theme: brand
      text: Get Started
      link: /guide/getting-started
    - theme: alt
      text: Excel Guide
      link: /guide/excel
    - theme: alt
      text: View on GitHub
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: Real-Time In-Cell Formulas
    details: Render structures directly into Excel cells with =SMILES2IMG(A2). Automatically scales with cell size and recalculates seamlessly.
  - icon: ✏️
    title: Double-Click In-Place Editing
    details: Double-click any rendered molecule or reaction in Excel, Word, or PowerPoint to launch the embedded Ketcher 2D editor.
  - icon: 🧪
    title: Lossless KET Reaction Canvases
    details: Preserves multi-step reaction pathways, arrows, plus signs, and chemical annotations losslessly in .xlsx, .docx, and .pptx documents.
  - icon: 🛡️
    title: 100% Local & Privacy-Safe
    details: Powered by embedded EPAM Indigo and Ketcher runtimes. Zero cloud calls, zero external API dependencies, 100% offline.
  - icon: 🎨
    title: Publication & Patent Quality
    details: ACS 1996 drawing specifications, high-contrast monochrome patent line art, dark theme palettes, and transparent backgrounds.
  - icon: 📦
    title: Unified Office Suite Installer
    details: One single PowerShell command automates setup, file unblocking, and registry configuration for Excel, Word, and PowerPoint.
---

## ⚡ Quick Install

Install the entire suite in seconds via PowerShell:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

> **Requirements:** Windows 10/11, Microsoft Office (Excel / Word / PowerPoint 2016–365, 32-bit or 64-bit), and WebView2 Runtime.
