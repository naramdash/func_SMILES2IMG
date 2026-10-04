# SMILES2IMG — Automatic Molecular Structure Renderer for Excel

**English** | [한국어](README.KR.md)

[![Platform](https://img.shields.io/badge/Platform-Windows%20Excel%20(32%2F64bit)-blue.svg)](#runtime-requirements)
[![Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4.svg)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-green.svg)](https://excel-dna.net/)
[![EPAM Indigo](https://img.shields.io/badge/EPAM%20Indigo-1.48.0--rc.1-orange.svg)](https://lifescience.opensource.epam.com/indigo/)
[![Single-File XLL](https://img.shields.io/badge/Deployment-100%25%20Standalone%20XLL-brightgreen.svg)](#-quick-start-installation)

> **An Excel Add-in that instantly converts SMILES chemical structure strings into crisp, in-cell molecular images.**  
> 100% local and offline execution without external web requests, cloud services, or browser runtimes.

---

## 📑 Table of Contents
- [🚀 Quick Start (Installation)](#-quick-start-installation)
- [💡 Usage & Real-Time IntelliSense](#-usage--real-time-intellisense)
- [📖 Function Syntax & Parameter Specification](#-function-syntax--parameter-specification)
- [🎨 Practical Recipes & Examples](#-practical-recipes--examples)
- [🔬 Chemical Structure Rendering Optimizations (ACS Standards)](#-chemical-structure-rendering-optimizations-acs-standards)
- [🏗️ Architecture & Internal Lifecycle](#️-architecture--internal-lifecycle)
- [📂 Project Directory Structure](#-project-directory-structure)
- [🛠️ Building & Distribution (For Developers)](#️-building--distribution-for-developers)
- [🧪 Automated Testing & Verification](#-automated-testing--verification)
- [❓ Frequently Asked Questions (FAQ)](#-frequently-asked-questions-faq)

---

## 🚀 Quick Start (Installation)

### Runtime requirements

| Component | Requirement |
| :--- | :--- |
| Windows | Windows 10 or 11; Windows 11 recommended |
| Excel (perpetual license) | Excel 2024 or later with native in-cell image support |
| Excel (Microsoft 365) | An updated Windows desktop version with **Place in Cell** available |
| Runtime | **.NET Framework 4.8 or later** |
| Excel architecture | Both **32-bit and 64-bit** packages are provided; choose the XLL matching Excel's bitness |

The current implementation requires native in-cell images. **Excel 2021, 2019, and 2016 are not supported.** Microsoft's [Place in Cell documentation](https://support.microsoft.com/en-us/excel/insert-picture-in-cell-in-excel) lists Microsoft 365 and Excel 2024 as supported versions. Check **Insert → Pictures → Place in Cell** in Excel to confirm the feature is available. The add-in runs only in Windows desktop Excel; Excel for Mac and Excel for the web cannot load this XLL.

These requirements are based on the current implementation and its dependencies; the project has not established a tested minimum Windows/Excel build across all versions. The **.NET SDK is needed only for development/building**, not for using the downloaded XLL.

### ⚡ Method A: One-Click Automatic Installer & Updater (Recommended)

Open PowerShell and run this **single command** to install, register, or update everything automatically:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

> **What the installer automates:**  
> 1. Detects your Excel architecture (32-bit vs. 64-bit)  
> 2. Downloads and updates the latest `.xll` in the official `%APPDATA%\Microsoft\AddIns` directory  
> 3. Unblocks the file (`Unblock-File`) to remove Windows security flags  
> 4. Registers the add-in in Excel options for persistent auto-loading  
>  
> 💡 **Future Updates:** Running this exact command again in the future automatically upgrades SMILES2IMG to the latest release!  
> 🗑️ **Uninstallation:** To uninstall anytime, simply run `.\uninstall.ps1` from the repository.

---

### 🖐️ Method B: Manual Step-by-Step Installation

#### Step 1: Check your Excel bitness (32-bit vs. 64-bit)
1. In Excel, go to **[File] → [Account] → [About Excel]**.
2. Note whether the first line indicates **`32-bit`** or **`64-bit`**.

#### Step 2: Prepare the Add-in (.xll) file
Download the single `.xll` file matching your Excel bitness from the [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest) page (recommended location: `Win+R` ➔ `%APPDATA%\Microsoft\AddIns`):
* **64-bit Excel:** [**`Smiles2Img-AddIn64-packed.xll`** (v1.0.0 Download)](https://github.com/naramdash/func_SMILES2IMG/releases/download/v1.0.0/Smiles2Img-AddIn64-packed.xll)
* **32-bit Excel:** [**`Smiles2Img-AddIn-packed.xll`** (v1.0.0 Download)](https://github.com/naramdash/func_SMILES2IMG/releases/download/v1.0.0/Smiles2Img-AddIn-packed.xll)

> [!TIP]
> **Unblock downloaded file (Essential):**  
> Web downloads receive an NTFS `Mark of the Web` tag, which causes Excel to block the add-in.  
> Right-click the downloaded `.xll` file → select **Properties** → check **Unblock** at the bottom of the General tab → click **OK**.

#### Step 3: Register as an Excel Add-in
1. In Excel, go to **[File] → [Options] → [Add-ins]**.
2. At the bottom, ensure **Manage: [Excel Add-ins]** is selected and click **[Go...]**.
3. Click **[Browse...]** and select your downloaded XLL file.
4. Verify that `Smiles2Img-AddIn` is checked in the list and click **[OK]**.

*(Once registered, the `=SMILES2IMG` function will automatically load every time Excel starts.)*

---

## 💡 Usage & Real-Time IntelliSense

Type `=SMILES2IMG(...)` in any worksheet cell to insert an in-cell molecular rendering:

```excel
=SMILES2IMG(A2)
```

```text
=SMILES2IMG(
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ SMILES2IMG(smiles, [background], [color], [transform])                                 │
│ A SMILES string or a cell reference containing one (e.g., A2, "CCO").                  │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

> [!TIP]
> • **Real-time formula tooltips (ExcelDna.IntelliSense):** As you type `=SMILES2IMG(`, Excel displays native-style tooltips highlighting the active argument in bold with clear English descriptions.  
> • **Cell sizing:** For optimal bond resolution and readability, increase the row height (e.g., 80–120 pt) and column width (e.g., 25–40).

---

## 📖 Function Syntax & Parameter Specification

```excel
=SMILES2IMG(smiles, [background], [color], [transform])
```

| Order | Argument | Type | Required | Default | Allowed Values & Description |
| :---: | :--- | :---: | :---: | :---: | :--- |
| **1** | **`smiles`** | String | **Required** | - | A valid SMILES string (1–2000 chars) or a cell reference (e.g., `A2`, `"CCO"`). |
| **2** | **`background`** | String | Optional | `"trans"` | **Background styling (case-insensitive):**<br>• `"trans"`, `"transparent"`, `"nobg"`: Transparent background (default, seamlessly blends with cell fills & zebra tables)<br>• `"white"`: Opaque white canvas background<br>• CSS color names: `"yellow"`, `"lightblue"`, `"lightgray"`, `"aliceblue"`, `"cornsilk"`, `"pink"`, etc. (140+ CSS standard names)<br>• Hex colors: `"#FFFF00"` or `"FFFF00"` (3- or 6-digit hex) |
| **3** | **`color`** | Boolean/Int | Optional | `TRUE` | **Atom color mode (academic/print style):**<br>• `TRUE` or `1`: Standard elemental colors (Oxygen=red, Nitrogen=blue, Sulfur=dark gold, etc.)<br>• `FALSE` or `0`: Grayscale / black-and-white mode |
| **4** | **`transform`** | String | Optional | `""` | **2D coordinate transforms (case-insensitive):**<br>• `flip`: Horizontal flip (matches textbook/reference ring orientation)<br>• `flipy`: Vertical flip<br>• `rot90`, `rot180`, `rot270`: Clockwise rotation by 90-degree increments |

> [!TIP]
> • **Skipping optional arguments:** To keep preceding arguments at their defaults (transparent background) while specifying later ones, use consecutive commas (e.g., `=SMILES2IMG(A2, , , "flip")`).  
> • **Smart fallback:** If a boolean `FALSE` is passed directly as the 2nd argument, the parser intelligently recognizes it as grayscale mode.

---

## 🎨 Practical Recipes & Examples

| Scenario | Example Formula | Note |
| :--- | :--- | :--- |
| **Default Rendering** | `=SMILES2IMG(A2)` | Cell reference (transparent background + standard colors) |
| **Direct SMILES String** | `=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O")` | Inline aspirin SMILES formula |
| **Opaque White Background** | `=SMILES2IMG(A2, "white")` | Explicit white canvas background |
| **CSS Color Background** | `=SMILES2IMG(A2, "yellow")`<br>`=SMILES2IMG(A2, "lightblue")` | Human-readable CSS color names |
| **Custom Hex Background** | `=SMILES2IMG(A2, "#FFFF00")` | Custom hex color background (`#` is optional) |
| **ACS Publication Grayscale** | `=SMILES2IMG(A2, , FALSE)` | Black & white publication rendering |
| **Textbook Layout Flip** | `=SMILES2IMG(A2, , , "flip")` | Flips horizontal orientation (e.g., nicotine, thiamine) |
| **White Canvas + B/W + Flip** | `=SMILES2IMG(A2, "white", FALSE, "flip")` | White background + monochrome + horizontal flip |
| **90° Rotation** | `=SMILES2IMG(A2, , , "rot90")` | Rotates long chains to fit cell aspect ratios |

---

## 🔬 Chemical Structure Rendering Optimizations (ACS Standards)

To satisfy the aesthetic and clarity standards of academic chemistry publications (ACS, IUPAC), the following automated enhancements are applied:

```mermaid
flowchart LR
    A["SMILES Input"] --> B["dearomatize<br>(Kekulé Form)"]
    B --> C["Smart Horizontal Layout<br>& Hetero-Methyl Labels"]
    C --> D["Preserve Ring-Junction<br>Chiral Hydrogens"]
    E["Insert Native In-Cell Image"]
    D --> F["Dynamic Resolution<br>Vector-grade PNG"]
    F --> E
```

1. **Dynamic Resolution & Vector-Grade Scaling:**  
   Replaces rigid canvas constraints with dynamic dimensioning based on bond length (80 px), margins (30 px), and relative bond thickness (1.5). Large macromolecular structures and simple diatomics maintain proportional bond weights and consistent label sizes.
2. **Kekulé Form Normalization (`dearomatize`):**  
   Converts ambiguous circular aromatic notations into crisp alternating double bonds preferred by organic chemists (`aromaticity-model = generic`).
3. **Selective Ring-Junction Chiral Hydrogen Preservation:**  
   While standard skeletal formulas hide carbon-bound hydrogens, bridgehead and ring-junction stereocenters (e.g., aflatoxin, bergenin) retain explicit wedge/dash hydrogens to prevent visual ambiguity and bond overlaps.
4. **Heteroatom-Bound Methyl Groups & Small-Molecule Readability:**  
   - Ultra-small molecules ($\le$ 4 heavy atoms, e.g., methyl isocyanate `CN=C=O`, methanol `CO`, ethanol) automatically render explicit text formulas (`H₃C-N=C=O`, `H₃C-OH`).
   - Heteroatom-bound methyls (e.g., in nicotine, caffeine, pyrethrins) display explicit labels (`N-CH₃`, `H₃C-O-`), while aliphatic hydrocarbon backbones remain sleek skeletal lines.
5. **Smart Horizontal Layout (`smart-layout` + `layout-orientation: horizontal`):**  
   Aliphatic bridges between rings (such as in thiamine and nicotine) are leveled horizontally, perfectly complementing standard rectangular spreadsheet cells.
6. **Optimized Sulfur (S) Visibility:**  
   Rendered in dark goldenrod (`#A88013`) instead of glaring yellow (`#FFFF00`) for clear contrast on white backgrounds.
7. **Alpha-Channel Transparency & Hex Colors:**  
   Supports true 32-bit alpha transparency (`Alpha=0`) and custom hex color fills.

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

- **100% In-Memory Offline Execution:** Eliminates external network calls, Node.js processes, or embedded browser frameworks.
- **Cache Persistence & Workbook Portability:** Images are cached in an internal hidden sheet (`__SMILES2IMG`), ensuring instant restoration when workbooks are saved and reopened.

---

## 📂 Project Directory Structure

```text
func_SMILES2IMG/
├── dist/                               # Standalone packed XLL distribution artifacts
│   ├── x64/
│   │   ├── Smiles2Img-AddIn64-packed.xll   # 64-bit self-contained Excel Add-in (~8.6 MB)
│   │   └── THIRD_PARTY_NOTICES.md
│   └── x86/
│       ├── Smiles2Img-AddIn-packed.xll     # 32-bit self-contained Excel Add-in (~8.7 MB)
│       └── THIRD_PARTY_NOTICES.md
├── Rendering/                          # Chemistry rendering & image pipeline
│   ├── MoleculeRenderer.cs             # Indigo rendering logic & ACS optimizations
│   ├── NativeImageWorkbook.cs          # OpenXML native image workbook support
│   └── NativeIndigo.cs                 # Embedded unmanaged C++ DLL extraction loader
├── assets/                             # Workbook templates & static resources
│   └── native-image-template.xlsx
├── Properties/
│   └── AssemblyInfo.cs
├── tests/                              # Automated test suites
│   ├── Excel-Smoke.ps1                 # Full COM automation E2E smoke tests
│   ├── Program.cs                      # Headless unit test runner
│   └── Smiles2Img.Tests.csproj
├── AddIn.cs                            # Excel-DNA initialization & diagnostic ribbon UI
├── CellImages.cs                       # In-cell image cache & ExcelReference bindings
├── CellImageUpdates.cs                 # Asynchronous image insertion macro queue
├── Functions.cs                        # =SMILES2IMG UDF & IntelliSense definitions
├── Smiles2Img-AddIn.dna                # Excel-DNA manifest configuration
├── Smiles2Img.csproj                   # MSBuild project file (.NET 4.8)
├── THIRD_PARTY_NOTICES.md              # Open-source license acknowledgments
├── PLAN.md                             # Architectural roadmap & optimization log
├── README.KR.md                        # Korean documentation
└── README.md                           # Main English documentation (this file)
```

---

## 🛠️ Building & Distribution (For Developers)

### Prerequisites
* Windows 10/11
* [.NET SDK 8.0 or newer](https://dotnet.microsoft.com/) (targeting .NET Framework 4.8)

### Build Commands
Run from the repository root:

```powershell
# Restore dependencies
dotnet restore --tl:off

# Build Release and pack standalone XLLs
dotnet build -c Release --no-restore --tl:off
```

The resulting packed XLLs in `dist/x64/` and `dist/x86/` contain all dependencies (ExcelDna, IntelliSense, Indigo C++ runtimes) within a single self-extracting archive.

---

## 🧪 Automated Testing & Verification

### 1. Headless Unit Tests (Excel not required)
Verifies molecule rendering, chiral hydrogen retention, transforms, and alpha transparency in headless console mode:

```powershell
dotnet build tests/Smiles2Img.Tests.csproj -c Release --no-restore --tl:off
& tests/bin/Release/net48/Smiles2Img.Tests.exe
```

### 2. Excel COM E2E Automation Smoke Tests (Runs live Excel)
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

---

## ❓ Frequently Asked Questions (FAQ)

* **Q. The formula returns `#VALUE!`.**
  * Check that the referenced cell contains a valid SMILES string.
  * If the target cell is part of **merged cells**, Excel formulas may point to an empty sub-cell, triggering `#VALUE!`.
* **Q. Excel blocks or refuses to load the XLL.**
  * Right-click the `.xll` file, open **Properties**, check **Unblock** at the bottom, and click **OK**.
* **Q. The molecular image looks too small or shrunk.**
  * Images scale automatically to preserve aspect ratios within cells. Increase the cell's row height (e.g., 100 pt) and column width (e.g., 30) for larger renderings.
* **Q. How do I inspect diagnostics and logs?**
  * From the Excel ribbon, click **[SMILES] → [Show diagnostics]** to view live rendering operations, cache statistics, and error logs.
