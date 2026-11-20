# Introduction

**`func_SMILES2IMG`** is a free, open-source chemical structure and 3D macromolecule add-in suite for **Microsoft Excel, Word, and PowerPoint**.

Write a compound name, CAS number, or SMILES directly in an Excel cell to render 2D structures and compute molecular properties across hundreds of rows—or double-click any structure in Word and PowerPoint to edit in place with the built-in **Ketcher 2D** sketcher and **Mol\* 3D** protein viewer.

## What You Can Do Across Microsoft Office

| Office App | Highlights | Guide |
| :--- | :--- | :--- |
| **Microsoft Excel** | In-cell structure formulas (`=SMILES2IMG`, `=NAME2IMG`), 40 property/Lipinski/SAR functions, `.sdf` import/export, double-click 2D editing | [Excel Cookbook](/excel/) |
| **Microsoft Word** | Inline structures in prose and compound tables, aspect-ratio-locked double-click re-editing, 1-click ACS B&W batch styling, vector SVG copy | [Word Tutorial](/word/) |
| **Microsoft PowerPoint** | Transparent white-ink Dark Mode structures, 0.1s textbox conversion, interactive 360°-rotatable 3D protein models (`.glb`) | [PowerPoint Gallery](/powerpoint/) |
| **Shared Editors** | Built-in offline **Ketcher 2D** sketcher (`Ctrl+Shift+K`) and **Mol\* 3D** macromolecular viewer across all three apps | [Editor Guides](/editor/) |

![func_SMILES2IMG running in Microsoft Excel](/excel-showcase.png)

## Why `func_SMILES2IMG`?

| Feature | **func_SMILES2IMG** | **ChemDraw / JChem for Office** | **Chem4Word** |
| :--- | :--- | :--- | :--- |
| **License & Cost** | **Free & Open-Source (MIT)** | Paid commercial subscription | Free & open-source |
| **Supported Office Apps** | **Excel + Word + PowerPoint** | Excel + Word + PowerPoint | Word only |
| **Excel Structure Cells** | **Native in-cell formulas (`=SMILES2IMG`)** & `.FLOAT` shapes (no sheet conversion needed) | Requires worksheet conversion or proprietary objects | N/A |
| **3D Macromolecules** | **Built-in Mol\* (rotatable `.glb` 3D models & 2D PNG)** | Requires separate 3D software | Not supported |
| **Installation Rights** | **No Admin rights needed** (10-second per-user setup) | Requires Admin installer & license activation | Requires setup wizard |

## Local-First Privacy & Cross-Platform Sharing {#privacy}

### Local Execution vs. Optional Network Calls
Core 2D rendering (**EPAM Indigo**), 2D sketching (**EPAM Ketcher**), 3D viewer runtime (**Mol\***), all physicochemical/Lipinski/SAR calculations, `.sdf` I/O, Batch Style, and SVG copy run **locally on your PC** without uploading your structures.

Internet access is used **only when you explicitly invoke** one of these optional features:
1. **PubChem Name / CAS Lookup:** Resolving an uncached compound name or CAS number (`=NAME2IMG`, `=NAME2SMILES`, `=SMILES2NAME`, or Substance Name in `Molecule Info`) queries PubChem once and caches the result locally on disk.
2. **RCSB PDB Fetch in Mol\*:** Downloading a public 3D structure by 4-letter PDB code (`1CRN`, `6LU7`, etc.) fetches from RCSB PDB / PDBe (opening local `.pdb`/`.cif` files works offline).
3. **Updates & Online Installer:** Clicking **`[Check for Updates]`** in `About` or running `install.ps1` connects to GitHub Releases.

### Viewing Documents Without the Add-In (Mac & Colleagues)
* **Word (`.docx`) & PowerPoint (`.pptx`):** 2D structures are saved as standard high-DPI PNG images and 3D proteins as standard `.glb` 3D models—anyone on Windows, **macOS**, iPad, or Office for the Web can view, print, and export to PDF without installing the add-in.
* **Excel (`.xlsx`):** Cached images and calculated values remain visible on other PCs; before sharing a final sheet externally where recalculation (<kbd>F9</kbd>) may occur, paste formula columns as values or export to PDF.

## System Requirements

| Component | Requirement | Notes |
| :--- | :--- | :--- |
| **OS** | Windows 10 / Windows 11 | Supports both 32-bit and 64-bit systems |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | Desktop installations |
| **Runtime** | .NET Framework 4.8 & Edge WebView2 | Pre-installed on Windows 10 & Windows 11 |

## Next Steps

- [**Installation (10-Second Setup)**](/guide/install) — Install across Excel, Word, and PowerPoint in one PowerShell line—no admin rights required.
- [**Getting Started (First Molecule)**](/guide/getting-started) — Render your first structure in Excel, Word, or PowerPoint and download the 8-tab master workbook.
