# func_SMILES2IMG — Chemical Structures & 3D Proteins inside Excel, Word & PowerPoint

[![Platform](https://img.shields.io/badge/Platform-Windows%20Office%20(Excel%20%7C%20Word%20%7C%20PowerPoint)-blue.svg)](https://naramdash.github.io/func_SMILES2IMG/guide/getting-started)
[![License](https://img.shields.io/badge/License-MIT-success.svg)](LICENSE)
[![Docs EN](https://img.shields.io/badge/Docs-English%20Guide-informational.svg)](https://naramdash.github.io/func_SMILES2IMG/)
[![Docs KO](https://img.shields.io/badge/문서-🇰🇷%20한국어%20가이드-0f766e.svg)](https://naramdash.github.io/func_SMILES2IMG/ko/)
[![Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4.svg)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-green.svg)](https://excel-dna.net/)
[![EPAM Indigo](https://img.shields.io/badge/EPAM%20Indigo-1.49.0--rc.2-orange.svg)](https://lifescience.opensource.epam.com/indigo/)
[![EPAM Ketcher](https://img.shields.io/badge/EPAM%20Ketcher-v3.18.0-blueviolet.svg)](https://lifescience.opensource.epam.com/ketcher/)
[![Mol* Viewer](https://img.shields.io/badge/Mol*--Viewer-v5.13.1-teal.svg)](https://molstar.org/)

> **Stop copy-pasting chemical structures into Office.**  
> Type a compound name or SMILES to auto-render 2D chemical structures, molecular weights, and Lipinski Rule of 5 across hundreds of Excel rows. Double-click any cell or picture to sketch in 2D with embedded EPAM Ketcher, or rotate 3D proteins with Mol\* — **100% free, local, and offline (no admin rights required).**

<p align="center">
  <img src="assets/hero-demo.gif" alt="func_SMILES2IMG Demo — Auto-render in Excel and double-click to edit in Ketcher 2D" width="820" />
</p>

---

## 📖 User Documentation & Interactive Web Tools

* 🇰🇷 **[한국어 공식 가이드 및 1분 시작하기](https://naramdash.github.io/func_SMILES2IMG/ko/)** · **[🛠️ 웹 수식 생성기 체험](https://naramdash.github.io/func_SMILES2IMG/ko/tools/formula-builder)**
* 🇺🇸 **[English Documentation & 1-Min Quickstart](https://naramdash.github.io/func_SMILES2IMG/)** · **[🛠️ Interactive Formula Generator](https://naramdash.github.io/func_SMILES2IMG/tools/formula-builder)**

---

## ⚡ 10-Second Quick Install (Excel · Word · PowerPoint)

Open **PowerShell** and run this single command to automatically install, configure security, and register the add-in across Microsoft Office (Excel, Word, PowerPoint):

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

> **What the installer automates:**  
> 1. Detects your Office bitness (32-bit vs. 64-bit) and installed Office applications (Excel, Word, PowerPoint)  
> 2. Downloads and places `func_SMILES2IMG` (Excel XLL, Word COM DLL, PowerPoint COM DLL) in `%APPDATA%\Microsoft\AddIns\Smiles2Img`  
> 3. Unblocks files (`Unblock-File`) to remove Windows security flags automatically  
> 4. Registers Excel, Word, and PowerPoint add-ins in HKCU registry (no admin rights required, loads automatically on start)  
> 💡 **Future Updates:** Running this exact command again in the future automatically upgrades to the latest release!  
> 🗑️ **Uninstallation:** Run `irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex` anytime to cleanly remove all files and registry entries.

---

## ✨ Key Features

* ⚡ **Real-Time In-Cell Chemical Rendering:** `=SMILES2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`, and `=NAME2IMG(A2)` render native cell pictures in Microsoft Excel that scale with row/column sizing and recalculate dynamically, with substructure highlighting (`"hl:c1ccccc1"`).
* 🔬 **Comprehensive Cheminformatics & Drug-Likeness Formulas:** Direct cell formulas `=SMILES2NAME`, `=NAME2SMILES`, `=SMILES2MW`, `=SMILES2EXACTMASS`, `=SMILES2FORMULA`, `=SMILES2INCHI`, `=SMILES2INCHIKEY`, `=SMILES2LOGP`, `=SMILES2TPSA`, `=SMILES2HBD`, `=SMILES2HBA`, `=SMILES2ROTBONDS`, `=SMILES2LIPINSKI`, `=SMILES.MATCH`, and `=SMILES.SIMILARITY`.
* 📋 **One-Click Molecule Info, Vector SVG Copy & Batch Style:** Right-click context submenu and Ribbon tools across Excel, Word, and PowerPoint to copy Substance Name, Formula, MW, SMILES, InChI, InChIKey, or crisp **Vector SVG (`Copy SVG`)**, normalize all structures in one click (**`Batch Style`**), and import/export `.sdf` files in Excel (**`SDF File`**).
* ✏️ **Embedded Ketcher 2D Visual Sketcher:** 100% offline molecular editor with warm WebView2 instance pooling (`<0.05s` instant reopen) supporting atoms, bonds, rings, charges, and multi-step reaction schemes.
* 🧬 **Embedded Mol\* 3D Macromolecular Viewer:** Visualize proteins, nucleic acids, and complexes directly across Excel, PowerPoint, and Word. Download PDB structures by ID, insert rotatable 3D GLB models or 2D publication images, and re-edit anytime with double-click.
* 🔄 **Lossless Document Persistence:** KET coordinates, chemical annotations, and Mol* 3D session recipes (`.molj`) are stored losslessly inside document XML metadata (`CustomXMLParts`), allowing seamless re-editing when sharing files with collaborators.
* 🛡️ **100% Local & Privacy-Preserving:** Powered by embedded EPAM Indigo, EPAM Ketcher, and Mol* runtimes. Zero cloud calls, zero data leakage, works completely offline.

---

## 🏗️ Architecture & Internal Lifecycle

```mermaid
sequenceDiagram
    autonumber
    actor User as Excel User
    participant F as Functions.cs
    participant C as CellImages.cs
    participant Cache as Hidden Cache Sheet (__SMILES2IMG)
    participant R as MoleculeRenderer (Indigo C++)
    participant U as CellImageUpdates.cs (Async Macro)

    User->>F: Enter formula =SMILES2IMG(A2, ...)
    F->>F: Parse arguments (Background, Color, Transform)
    F->>C: Query cell image reference
    C->>Cache: Lookup hash key via Excel MATCH
    alt Cache Hit
        Cache-->>F: Return existing cell ExcelReference (0ms)
        F-->>User: Immediate image display
    else Cache Miss
        C->>R: In-memory high-res PNG render
        R-->>C: Generated PNG byte stream
        C->>U: Enqueue async update & return temporary #N/A
        U->>Cache: Insert native in-cell picture into hidden sheet
        U->>F: Trigger worksheet recalculation
        F-->>User: Rendered molecule displayed in cell
    end
```

* **100% In-Memory Offline Execution:** Eliminates external network calls, Node.js processes, or cloud API frameworks.
* **Cache Persistence & Workbook Portability:** Images are cached in an internal hidden sheet (`__SMILES2IMG`), ensuring instant restoration when workbooks are saved and reopened.
* **Aspect-Ratio Preserving Shape Updates:** Word and PowerPoint add-ins replace structure images in-place without altering coordinates, table cell widths, or margins.

---

## 📂 Project Directory Structure

```text
func_SMILES2IMG/
├── Smiles2Img.sln                      # Visual Studio / .NET unified solution file
├── Directory.Build.props               # Shared version, TFM (net48), and compiler configuration
├── package.json                        # VitePress documentation dependencies and scripts
├── .github/
│   └── workflows/
│       └── deploy-docs.yml             # GitHub Actions workflow for building & deploying docs to GitHub Pages
├── docs/                               # VitePress documentation source
│   ├── .vitepress/
│   │   └── config.mts                  # VitePress site configuration & i18n routing
│   ├── public/                         # Static assets (images, icons)
│   ├── guide/                          # English guides (getting started, excel, word, powerpoint, ketcher, molstar, faq)
│   ├── reference/                      # English reference (formulas & parameter specs)
│   └── ko/                             # Korean guides & reference (ko/guide/*, ko/reference/*)
├── src/
│   ├── Smiles2Img.Core/                # Shared rendering, Ketcher 2D, Mol* 3D & notification core (net48)
│   │   ├── Rendering/                  # EPAM Indigo 1.49.0 wrapper, ACS standards, color/option parsers, PubChem resolver
│   │   ├── Editor/                     # Ketcher v3.18.0 & Mol* v5.13.1 forms, session managers, StructureMetadata, CopiedToast
│   │   ├── Notifications/              # Non-intrusive floating badges and notification service
│   │   └── COM/                        # Shared Office COM interfaces (IDTExtensibility2, IRibbonExtensibility)
│   ├── Smiles2Img.Excel/               # Excel-DNA standalone XLL add-in (net48)
│   │   ├── Functions.cs                # =SMILES2IMG, =KET2IMG, and property calculation UDFs (=SMILES2NAME, =SMILES2MW, etc.)
│   │   ├── FloatFunctions.cs           # =SMILES2IMG.FLOAT & =KET2IMG.FLOAT floating UDFs (Excel 2016-2021)
│   │   ├── CellImages.cs               # In-cell image cache manager
│   │   ├── FloatingPictures.cs         # Cell-anchored floating shape manager
│   │   ├── SmilesRibbon.cs             # [SMILES] ribbon tab & right-click context menu handlers (Copy Molecule Info submenu)
│   │   ├── KetcherEditor.cs            # Ketcher commands, quick paste, molecule info & double-click hooks
│   │   └── KetcherEditor.ShapesAndCells.cs # KET canvas shape & formula cell extraction/insertion helpers
│   │   └── ExcelMolstarManager.cs      # Mol* 3D viewer integration, floating GLB/PNG insertion & session re-editing
│   ├── Smiles2Img.Word/                # Word COM add-in DLL (net48)
│   │   ├── WordAddIn.cs                # IDTExtensibility2 + IRibbonExtensibility + context menu handlers
│   │   └── WordStructureManager.cs     # Inline structure images, Mol* 3D models, replacement, quick paste & double-click
│   └── Smiles2Img.PowerPoint/          # PowerPoint COM add-in DLL (net48)
│       ├── PowerPointAddIn.cs          # IDTExtensibility2 + IRibbonExtensibility + context menu + double-click hook
│       └── PowerPointStructureManager.cs # Slide structure shape insertion, Mol* 3D models, replacement, quick paste & metadata
├── dist/                               # Distribution output directory
│   ├── x64/
│   │   ├── Smiles2Img-AddIn64-packed.xll # 64-bit standalone Excel XLL
│   │   ├── Word/                       # 64-bit Word COM add-in binaries
│   │   ├── PowerPoint/                 # 64-bit PowerPoint COM add-in binaries
│   │   └── func_SMILES2IMG-x64.zip     # 64-bit unified func_SMILES2IMG package
│   └── x86/
│       ├── Smiles2Img-AddIn-packed.xll   # 32-bit standalone Excel XLL
│       ├── Word/                       # 32-bit Word COM add-in binaries
│       ├── PowerPoint/                 # 32-bit PowerPoint COM add-in binaries
│       └── func_SMILES2IMG-x86.zip     # 32-bit unified func_SMILES2IMG package
├── assets/                             # Visual showcase assets & documentation screenshots
├── tests/                              # Automated test suites
│   ├── Excel-Smoke.ps1                 # Full COM automation E2E smoke tests
│   ├── Excel-365-FeatureTests.ps1      # Live Excel 365 new styles & full-option tests
│   ├── Excel-DeepStressTests.ps1       # Bulk rendering & edge-case stress test
│   ├── Program.cs                      # Unit tests + live WebView2/Ketcher v3.18.0 E2E round-trip runner
│   └── Smiles2Img.Tests.csproj
├── scripts/
│   ├── Test-OfficeE2E.ps1              # Unified Word, PowerPoint, and Excel COM Add-in E2E test suite
│   ├── Package-Release.ps1             # Release archive packaging script
│   └── Generate-Showcase-Workbook.ps1  # Showcase workbook generator
├── install.ps1                         # Unified Office Suite one-click installer
├── uninstall.ps1                       # Unified Office Suite uninstaller
├── THIRD_PARTY_NOTICES.md              # Open-source license acknowledgments
└── README.md                           # Main developer documentation (this file)
```

---

## 🛠️ Building & Distribution (For Developers)

### Prerequisites
* Windows 10 / 11
* [.NET SDK 8.0 or newer](https://dotnet.microsoft.com/) (targeting .NET Framework 4.8)
* [Node.js 18+ & npm](https://nodejs.org/) (for building the VitePress documentation)

### Full Solution Build Command
Run from the repository root:

```powershell
# Restore dependencies & build entire solution (packs Excel XLL and compiles Word/PowerPoint add-ins)
dotnet build Smiles2Img.sln -c Release

# Generate release ZIP archives
powershell -ExecutionPolicy Bypass -File scripts\Package-Release.ps1
```

The resulting packed XLLs in `dist/x64/` and `dist/x86/` contain all dependencies (ExcelDna, IntelliSense, Indigo C++ runtimes) within a single self-extracting archive.

### Documentation Build Command
To develop or build the VitePress documentation locally:

```powershell
# Start local development server with hot reload
npm run docs:dev

# Build static documentation for production
npm run docs:build

# Preview production build locally
npm run docs:preview
```

---

## 🧪 Automated Testing & Verification

### 1. Core Unit Tests + Live WebView2 Ketcher E2E Tests (Office not required)
Verifies molecule rendering, chiral hydrogen retention, transforms, alpha transparency, molecule properties (`Formula`, `MW`, `Exact Mass`), Office Ribbon XML reflection, and live WebView2 + EPAM Ketcher v3.18.0 round-trip (`SMILES` $\leftrightarrow$ `KET` + Transparent/White PNG export):

```powershell
dotnet build tests/Smiles2Img.Tests.csproj -c Release --no-restore --tl:off
& tests/bin/Release/net48/Smiles2Img.Tests.exe
```

### 2. Unified Office Suite COM E2E Tests (Word + PowerPoint + Excel)
Launches live instances of **Microsoft Word, PowerPoint, and Excel** to verify COM add-in registration, `KET` + `SMILES` metadata round-tripping (`AlternativeText`, `Title`, `Tags`), aspect-ratio preserving in-place updates, and formula evaluation:

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File scripts/Test-OfficeE2E.ps1
```

### 3. Excel COM E2E Automation Smoke Tests (Runs live Excel)
Automates an actual background Excel instance to test full lifecycle reliability:

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-Smoke.ps1
```
* XLL registration and formula discovery
* In-cell picture embedding and formula persistence
* Real-time image replacement upon SMILES modification (zero floating shapes)
* Cache hit instantaneous reuse across multiple cells
* Workbook saving, reopening, and cache restoration
* Automatic orphaned cache cleanup when formulas are deleted

### 4. Excel 365 Real-World Feature & Bulk Stress Tests
Validates modern delimiters, ultimate full-option formulas, and 20+ concurrent complex molecules:

```powershell
# Live Excel 365 feature & full-option suite
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-365-FeatureTests.ps1

# 20-molecule bulk rendering & stress suite
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-DeepStressTests.ps1
```

---

## 📄 License & Third-Party Notices

* Licensed under the [MIT License](LICENSE).
* See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for licenses and copyright notices of embedded third-party components (EPAM Indigo, EPAM Ketcher, Excel-DNA, Microsoft Edge WebView2).
