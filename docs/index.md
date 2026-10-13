---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "Free & Open-Source Office Chemical Add-in"
  tagline: "Free, open-source 2D chemical structure rendering, real-time in-cell formulas, seamless double-click editing, and lossless KET reaction canvases across Excel, Word, and PowerPoint — 100% offline & local."
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

Install the entire suite in seconds via **PowerShell**:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 Automatic Upgrade
Running this exact command anytime upgrades your installation to the latest release!
:::

---

## 🖼️ Application Showcase

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel](/guide/excel) | [📘 Microsoft Word](/guide/word) | [📙 Microsoft PowerPoint](/guide/powerpoint) |
| :---: | :---: | :---: |
| Native in-cell formulas (`=SMILES2IMG`), cell resizing, sorting & filtering | Inline shapes flowing with paragraphs, table cell preservation | Center slide auto-placement, dark theme transparent background |
