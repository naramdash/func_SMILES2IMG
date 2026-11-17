# func_SMILES2IMG — Chemical Structures & 3D Proteins inside Excel, Word & PowerPoint

[![Platform](https://img.shields.io/badge/Platform-Windows%20Office%20(Excel%20%7C%20Word%20%7C%20PowerPoint)-blue.svg)](https://naramdash.github.io/func_SMILES2IMG/guide/getting-started)
[![License](https://img.shields.io/badge/License-MIT-success.svg)](LICENSE)
[![Docs EN](https://img.shields.io/badge/Docs-English%20Guide-informational.svg)](https://naramdash.github.io/func_SMILES2IMG/)
[![Docs KO](https://img.shields.io/badge/문서-🇰🇷%20한국어%20가이드-0f766e.svg)](https://naramdash.github.io/func_SMILES2IMG/ko/)
[![Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4.svg)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-green.svg)](https://excel-dna.net/)
[![EPAM Indigo](https://img.shields.io/badge/EPAM%20Indigo-1.49.0-orange.svg)](https://lifescience.opensource.epam.com/indigo/)
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
* 📊 **[📥 Download 8-Tab Master Interactive Excel Showcase (.xlsx)](https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/SMILES2IMG_Comprehensive_Showcase.xlsx)** — 8 category tabs featuring live in-cell structures, 18-column Lipinski Ro5 & 3D complexity descriptor matrix (MW, LogP, TPSA, Rings, Fsp3, Chiral), substructure matching, styling matrix, and reaction canvases.

---

## ⚡ 10-Second Quick Install (Excel · Word · PowerPoint)

Open **PowerShell** and run this single command to automatically install, configure security, and register the add-in across Microsoft Office (Excel, Word, PowerPoint):

```powershell
# Stable release (v1.7.2)
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex

# Preview release (v1.8.0-preview.2 — Unified Flyout Context Menus, Cross-Host Parity, 3D Complexity Descriptors & 12-Col SDF Import)
$v='v1.8.0-preview.2'; irm https://github.com/naramdash/func_SMILES2IMG/releases/download/v1.8.0-preview.2/install.ps1 | iex
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
* 🔬 **Comprehensive Cheminformatics & Drug-Likeness Formulas:** Direct cell formulas `=SMILES2NAME`, `=NAME2SMILES`, `=SMILES2MW`, `=SMILES2EXACTMASS`, `=SMILES2FORMULA`, `=SMILES2INCHI`, `=SMILES2INCHIKEY`, `=SMILES2LOGP`, `=SMILES2TPSA`, `=SMILES2HBD`, `=SMILES2HBA`, `=SMILES2ROTBONDS`, `=SMILES2LIPINSKI`, `=SMILES2HEAVY`, `=SMILES2RINGS`, `=SMILES2AROMATIC`, `=SMILES2FSP3`, `=SMILES2CHIRAL`, `=SMILES.MATCH`, and `=SMILES.SIMILARITY`.
* 📋 **One-Click Molecule Info, Vector SVG Copy & Batch Style:** Right-click context submenu and Ribbon tools across Excel, Word, and PowerPoint to copy Substance Name, Formula, MW, SMILES, InChI, InChIKey, or crisp **Vector SVG (`Copy SVG`)**, normalize all structures in one click (**`Batch Style`**), and import/export 12-column `.sdf` chemical database tables in Excel (**`SDF File`**).
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
├── .editorconfig & .vscode/            # Enforces UTF-8 (no BOM) & CRLF across editors
├── .github/workflows/                  # GitHub Actions workflows
│   ├── ci.yml                          # Release build + 24 Unit/WebView2 E2E test suites
│   ├── release.yml                     # Release packaging, installer encoding check & GitHub Release
│   └── deploy-docs.yml                 # Docs link audit & GitHub Pages deployment
├── docs/                               # VitePress documentation source
│   ├── .vitepress/                     # VitePress config, i18n routing & interactive Vue widgets
│   ├── public/                         # Static assets (images, icons, sample SDF, showcase XLSX)
│   ├── guide/                          # English guides (getting started, excel, word, powerpoint, ketcher, molstar, faq)
│   ├── reference/                      # English reference (formulas & parameter specs)
│   └── ko/                             # Korean guides & reference (ko/guide/*, ko/reference/*)
├── src/
│   ├── Smiles2Img.Core/                # Shared rendering, Ketcher 2D, Mol* 3D, UI & COM core (net48)
│   │   ├── Rendering/                  # EPAM Indigo 1.49.0 wrapper (MoleculeRenderer, MoleculeInfo, SdfManager, PubChemResolver)
│   │   ├── Editor/                     # Ketcher v3.18.0 & Mol* v5.13.1 (WebView2 1.0.4314.35), RibbonXmlBuilder, WebView2HostHelper, MoleculeActionHelper
│   │   ├── Notifications/              # Non-intrusive floating badges and notification service
│   │   └── COM/                        # Shared Office COM interfaces (IDTExtensibility2, IRibbonExtensibility)
│   ├── Smiles2Img.Excel/               # Excel-DNA standalone XLL add-in (net48)
│   │   ├── Functions.cs                # 36 in-cell UDFs (=SMILES2IMG, =SMILES2MW, =SMILES2FSP3, =SMILES.MATCH, etc.)
│   │   ├── FloatFunctions.cs           # 4 floating shape UDFs (.FLOAT for Excel 2016-2021)
│   │   ├── CellImages.cs               # In-cell image cache manager
│   │   ├── FloatingPictures.cs         # Cell-anchored floating shape manager
│   │   ├── SmilesRibbon.cs             # [SMILES] ribbon tab & right-click context menu callbacks
│   │   ├── KetcherEditor.cs            # Ketcher commands, quick paste, SDF I/O, molecule info & double-click hooks
│   │   ├── KetcherEditor.ShapesAndCells.cs # KET canvas shape & formula cell extraction/insertion helpers
│   │   └── ExcelMolstarManager.cs      # Mol* 3D viewer integration, floating GLB/PNG insertion & session re-editing
│   ├── Smiles2Img.Word/                # Word COM add-in DLL (net48)
│   │   ├── WordAddIn.cs                # IDTExtensibility2 + IRibbonExtensibility + context menu callbacks
│   │   ├── WordStructureManager.cs     # 2D inline/floating structure images, Batch Style, quick paste & double-click
│   │   └── WordStructureManager.Molstar.cs # Partial class handling Word Mol* 3D models & CustomXMLParts persistence
│   └── Smiles2Img.PowerPoint/          # PowerPoint COM add-in DLL (net48)
│       ├── PowerPointAddIn.cs          # IDTExtensibility2 + IRibbonExtensibility + context menu + double-click hook
│       ├── PowerPointStructureManager.cs # 2D slide structure shape insertion, Batch Style, quick paste & metadata
│       └── PowerPointStructureManager.Molstar.cs # Partial class handling slide Mol* 3D models & persistence
├── dist/                               # Distribution output directory (generated by Release build / packaging)
│   ├── x64/                            # 64-bit standalone XLL, Word/PowerPoint DLLs & func_SMILES2IMG-x64.zip
│   └── x86/                            # 32-bit standalone XLL, Word/PowerPoint DLLs & func_SMILES2IMG-x86.zip
├── assets/                             # Visual showcase assets & documentation screenshots
├── tests/                              # Categorized Unit & E2E test suites
│   ├── Unit/                           # RenderingUnitTests, CheminformaticsUnitTests, UiAndRibbonUnitTests
│   ├── E2E/                            # WebView2EditorE2ETests (Ketcher/Mol*/Dialogs) & OfficeLiveE2ERunner
│   ├── TestAssert.cs                   # Shared test assertions & suite logging
│   ├── Program.cs                      # Test runner CLI entry point (supports --full)
│   ├── Excel-Smoke.ps1                 # Standalone Excel COM automation smoke tests
│   ├── Excel-365-FeatureTests.ps1      # Live Excel 365 feature & full-option tests
│   ├── Excel-DeepStressTests.ps1       # Bulk rendering & edge-case stress test
│   ├── packages.lock.json              # Reproducible NuGet restore lock file
│   └── Smiles2Img.Tests.csproj
├── scripts/
│   ├── Test-OfficeE2E.ps1              # Unified Live Word, PowerPoint, and Excel COM/XLL + UI E2E suite
│   ├── Package-Release.ps1             # Builds Release, strips UTF-8 BOM from installers & creates ZIP archives
│   ├── Generate-Showcase-Workbook.ps1  # 8-tab comprehensive showcase workbook generator
│   └── audit-docs-links.js             # Audits internal Markdown links & download assets across docs/
├── install.ps1                         # Unified Office Suite 1-click installer (clean-install/upgrade summary)
├── uninstall.ps1                       # Unified Office Suite uninstaller (removal summary report)
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

### 1. Category 1 Unit Tests (18 Suites) + Category 2 WebView2 / WinForms UI E2E (6 Suites)
Runs **24 test suites** (no Microsoft Office installation required): verifies 2D PNG/SVG rendering, chiral/stereochemical rules, transforms, molecule properties and complexity descriptors (`Formula`, `MW`, `Exact Mass`, `LogP`, `TPSA`, `Lipinski`, `Heavy Atoms`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`), Excel/Word/PowerPoint Ribbon & ContextMenu XML (`ValidateRibbonXml` + `VerifiedValidOfficeImageMsos`), WinForms dialogs (`MoleculeInfoDialog`, `AboutDialog`, `CopiedToast`), and live WebView2 Ketcher v3.18.0 + Mol\* v5.13.1 round-trip UI capture:

```powershell
dotnet run --project tests -c Release
```

### 2. Full End-to-End Test Suite Including Live Office COM/XLL Automation (25 Suites)
Executes all 24 Unit & WebView2 UI suites plus live Microsoft **Word (COM)**, **PowerPoint (COM)**, and **Excel (Live `.xll` + 21 UDFs + SDF Round-Trip + Modal UI popups)** automation:

```powershell
dotnet run --project tests -c Release -- --full

# Or run the live Office suite script directly after building:
powershell -NoProfile -STA -ExecutionPolicy Bypass -File scripts/Test-OfficeE2E.ps1
```

### 3. Additional Excel COM Automation Smoke & Stress Scripts
```powershell
# Standalone Excel lifecycle & cache restoration smoke tests
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-Smoke.ps1

# Live Excel 365 feature & full-option suite
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-365-FeatureTests.ps1

# 20-molecule bulk rendering & stress suite
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-DeepStressTests.ps1
```

### 4. Documentation Link Audit & Static Site Build
```powershell
node scripts/audit-docs-links.js
npm run docs:build
```

---

## 📄 License & Third-Party Notices

* Licensed under the [MIT License](LICENSE).
* See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for licenses and copyright notices of embedded third-party components (EPAM Indigo `1.49.0`, EPAM Ketcher `v3.18.0`, Mol\* `v5.13.1`, IUPAC InChI Toolkit `1.07.1`, Excel-DNA `1.9.0`, Microsoft Edge WebView2 `1.0.4314.35`).
