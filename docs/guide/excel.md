# SMILES for Microsoft Excel

**SMILES for Microsoft Excel** renders high-resolution 2D chemical structure graphics directly inside worksheet cells with instant in-cell formula recalculation, comprehensive ribbon commands, and seamless double-click editing.

---

## 🌟 Visual Feature Walkthrough

All core chemical operations in Excel are integrated natively into the Ribbon tab, right-click context menus, and worksheet formula engine:

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip Quick Start Checklist
1. **Draw in an Empty Cell:** Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>, draw your molecule or reaction in Ketcher, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
2. **Formula Batch Insertion:** Type `=SMILES2IMG(A2)` or `=KET2IMG(A2)` and drag down across hundreds of rows.
3. **Quick Insert from Clipboard:** Copy any SMILES/KET text and click **[SMILES] ➔ [Insert from Clipboard]**.
4. **Edit Anytime:** Double-click the cell to re-edit in Ketcher, or right-click to copy molecular properties (Formula, MW, Exact Mass).
:::

---

## 📑 Dedicated Office Ribbon: `[SMILES]` Tab

The custom `[SMILES]` Ribbon tab provides one-click access to all chemical structure operations without needing to remember formula syntax or keyboard shortcuts:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ Insert Structure (`Ctrl+Shift+K`):** Launches the offline Ketcher 2D editor. Draw structures, reactions, or functional groups, then press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert directly into the active cell.
* **🧬 3D Structure (Mol\*):** Opens the embedded Mol\* 3D Macromolecular Viewer. Download biological macromolecules by PDB ID or open local structure files, and insert interactive 3D models (`.glb`) or 2D images (`.png`) onto your worksheet.
* **⚡ Quick Insert:** Converts highlighted SMILES strings into rendered structures in 0.1s. If an empty cell is selected, it automatically reads the clipboard.
* **📋 Insert from Clipboard:** Parses clipboard text as SMILES or Ketcher KET and creates an in-cell chemical structure immediately.
* **🖼️ Copy SVG *(New in v1.7.0)*:** Renders the selected molecule as a scalable vector graphic (`SVG`) and copies it to the Windows Clipboard for lossless pasting into Adobe Illustrator, Inkscape, or vector design tools.
* **🎨 Batch Style *(New in v1.7.0)*:** Converts all chemical structures in the active worksheet to a unified style preset (**Standard Color (Transparent)**, **Publication B&W (ACS Monochrome)**, **White Background (Color)**, or **Dark Mode (White Ink)**) in one click.
* **📂 SDF File (`Import .sdf` / `Export .sdf`) *(New in v1.7.0)*:** Imports multi-molecule `.sdf` chemical databases directly into an Excel table with auto-generated `=SMILES2IMG(...)`, Formula, MW, and property columns, or exports worksheet molecules back to `.sdf`.
* **🔍 Settings & Diagnostics:** Opens the real-time engine health monitor, verifying Indigo engine loading, WebView2 bundle integrity, and formula recalculation logs.

---

## 🖱️ Redesigned Context Menu: `Chemical Structure`

Right-click anywhere in your worksheet to access contextual chemical actions cleanly separated from standard Excel paste options:

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Case A: Right-Clicking a Molecule Cell
When right-clicking a cell containing a chemical structure or formula:
* **✏️ Edit Structure in Ketcher:** Opens the Ketcher editor with the current molecule or reaction graph loaded for in-place modification.
* **🧬 3D Structure (Mol\*)...:** Launches the Mol\* 3D viewer. If the active cell contains a PDB ID (e.g., `1CRN`), Mol\* preloads the macromolecule automatically!
* **📋 Copy SMILES:** Copies clean SMILES notation to clipboard.
* **📋 Copy InChI:** Generates and copies standard IUPAC InChI string instantly.
* **📋 Copy Canvas Data (KET):** Copies Ketcher native JSON representation.
* **🖼️ Copy Vector Graphic (SVG):** Renders the structure as a crisp scalable vector graphic (`SVG`) and copies it to the clipboard for Adobe Illustrator or Inkscape.
* **🔬 Copy Molecule Info (Submenu):**
  * **Substance Name:** Copies substance common name (e.g., `Aspirin`, `Ethanol`).
  * **Chemical Formula:** Copies standard Hill-notation formula (e.g., `C9H8O4`).
  * **Molecular Weight (MW):** Copies IUPAC molecular weight in g/mol (e.g., `180.16 g/mol`).
  * **Copy SMILES / InChI / InChIKey:** Individual one-click copy options.
  * **All Summary (Text):** Copies full formatted summary to clipboard.
  * **Detailed Molecule Info Window...:** Opens the full chemical properties window with 2D SVG preview, **LogP / TPSA**, **Lipinski Rule of 5**, and complete property table.
  > ⚡ **Zero Modal Distraction:** Clicking any copy command displays a non-blocking floating toast badge (`✓ Copied {Item}: {Value}`) near your mouse cursor for 1.4 seconds without stealing focus or interrupting your work!

### Case B: Right-Clicking an Empty Cell or Structure Shape
* **✏️ Insert Structure in Ketcher:** Launches Ketcher to draw a new chemical structure and place it directly into the selected cell.
* **🧬 3D Structure (Mol\*)...:** Launches Mol\* to insert an interactive 3D model or 2D snapshot aligned to the active cell.
* **✏️ Edit 3D Structure (Mol\*)...:** (When right-clicking a Mol\* 3D model or picture) Re-opens Mol\* in the exact camera angle and representation style.
* **📋 Paste Structure:** Automatically detects chemical notation in the clipboard and inserts it as a rendered molecule.

---

## 🧪 Ketcher 2D Molecule & Reaction Editor

Double-clicking any chemical cell or pressing <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> opens the embedded, 100% offline Ketcher 2D chemical structure editor (with `<0.05s` warm WebView2 instance pooling for instant reopening):

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Key Ketcher Features
* **Live Molecular Property Banner:** As you draw bonds and functional groups, the top status bar calculates **Formula**, **Molecular Weight (MW)**, and **Exact Mass** in real time.
* **Background Transparency Toggle:** Switch between **[Background: White]** and **[Background: Transparent]** with a single click. Transparent mode generates alpha-channel PNGs that look stunning on dark spreadsheet themes and presentations.
* **Full Organic & Inorganic Toolkit:** Periodic table atom picker, standard rings (benzene, cyclopentane, cyclohexane), stereo bonds (wedges/hashes), reaction arrows, and catalyst text.
* **Keyboard Hotkeys:** Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to apply changes immediately to Excel, or <kbd>Esc</kbd> to discard.

---

## 🔬 Lossless Reaction Canvases: `=KET2IMG`

For multi-step organic reaction pathways involving reagents, catalysts, temperatures, yields, and plus signs (`+`), use the dedicated `=KET2IMG` formula:

```text
=KET2IMG(A2)
```

* **Lossless Coordinate Preservation:** Unlike SMILES which only stores molecular connectivity, Ketcher's native `KET` format stores 2D layout coordinates, reaction arrows, catalysts, and text notes with zero degradation.
* **Double-Click Re-editing:** Double-clicking the reaction cell re-opens Ketcher with the entire reaction scheme, allowing you to edit reaction conditions and yields in place.

---

## 📝 Functions: Native In-Cell vs. Floating Images

Depending on your Microsoft Excel version, two distinct formula modes are available:

### 1. `=SMILES2IMG`, `=INCHI2IMG`, `=KET2IMG`, & `=NAME2IMG`
* **Target:** **Microsoft 365** and **Excel 2024+**
* Embeds real **native in-cell pictures** directly into worksheet cell contents from SMILES, Ketcher KET, IUPAC InChI, or chemical names/CAS numbers (`=NAME2IMG("Aspirin")`, `=NAME2IMG("50-78-2")`).
* Moves, sorts, filters, and auto-scales with row height and column width.
* Can be referenced by standard Excel lookup functions, formulas, and conditional formatting rules.

### 2. `=SMILES2IMG.FLOAT`, `=INCHI2IMG.FLOAT`, `=KET2IMG.FLOAT`, & `=NAME2IMG.FLOAT`
* **Target:** **Excel 2016 / 2019 / 2021** (or when floating shapes are explicitly desired)
* Inserts floating picture shapes locked to cell coordinates (`xlMoveAndSize`).
* Recalculates automatically, replaces existing shapes without orphan clutter, and moves/resizes with cell boundaries.

### 3. Molecular Property, Drug-Likeness & SAR Functions
Calculate chemical characteristics and screen compound libraries directly in your worksheet:
* `=SMILES2NAME(cell)` / `=KET2NAME(cell)` / `=INCHI2NAME(cell)` — Resolves chemical substance name via PubChem/IUPAC.
* `=NAME2SMILES(name_or_cas)` — Resolves chemical common name or CAS Registry Number (`"50-78-2"`) to canonical SMILES.
* `=SMILES2MW(cell)` / `=KET2MW(cell)` / `=INCHI2MW(cell)` — Computes standard molecular weight in g/mol (0ms pure local).
* `=SMILES2EXACTMASS(cell)` / `=KET2EXACTMASS(cell)` / `=INCHI2EXACTMASS(cell)` — Computes monoisotopic exact mass in Da for HRMS/MS (0ms pure local).
* `=SMILES2FORMULA(cell)` / `=KET2FORMULA(cell)` / `=INCHI2FORMULA(cell)` — Generates Hill-notation formula (e.g. `C9H8O4`, 0ms pure local).
* `=SMILES2INCHI(cell)` / `=KET2INCHI(cell)` & `=INCHI2SMILES(cell)` — Bidirectional conversion between SMILES/KET and IUPAC InChI (0ms pure local).
* `=SMILES2INCHIKEY(cell)` / `=KET2INCHIKEY(cell)` / `=INCHI2INCHIKEY(cell)` — Computes 27-character unique InChIKey hash (0ms pure local).
* `=SMILES2LOGP(cell)` & `=SMILES2TPSA(cell)` — Computes Wildman-Crippen LogP and Topological Polar Surface Area (Å², 0ms pure local).
* `=SMILES2HBD(cell)`, `=SMILES2HBA(cell)`, `=SMILES2ROTBONDS(cell)` — Counts H-bond donors, H-bond acceptors, and rotatable single bonds (0ms pure local).
* `=SMILES2LIPINSKI(cell)` — Evaluates Lipinski's Rule of 5 (`"PASS"` or `"FAIL"`, 0ms pure local).
* `=SMILES.MATCH(target, query)` & `=SMILES.SIMILARITY(mol1, mol2, [metric])` — Substructure matching (`TRUE`/`FALSE`) and Tanimoto/Tversky 2D fingerprint similarity (`0.0` to `1.0`, 0ms pure local).

---

## 🎨 Chemical Styling & Rendering Options

Options can be combined freely using spaces, commas, semicolons, or pipe symbols:

| Option Keyword | Description | Example |
| :--- | :--- | :--- |
| `bw` / `monochrome` | High-contrast black-and-white line art (ACS / USPTO patent standard) | `=SMILES2IMG(A2, "", "bw")` |
| `colored` | Standard IUPAC element color-coded atoms | `=SMILES2IMG(A2, "", "colored")` |
| `hl:<smarts>` | Highlights atoms and bonds matching a SMILES/SMARTS substructure query | `=SMILES2IMG(A2, "trans", "hl:c1ccccc1")` |
| `num` / `numbers` | Displays atom index numbers (ideal for NMR/MS peak assignments) | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | Forces CIP stereochemistry labels (*R/S*, *E/Z*) | `=SMILES2IMG(A2, "", "stereo")` |
| `h` / `all-c` | Unfolds explicit hydrogens (`h`) or displays all carbon symbols (`all-c`) | `=SMILES2IMG(A2, "", "h all-c")` |
| `white` / `ink=#HEX` | White or custom hex ink color for dark backgrounds | `=SMILES2IMG(A2, "#1e1e1e", "white")` |
| `flip` / `flipv` | Flips structure horizontally or vertically | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` | Rotates molecule in 90-degree increments | `=SMILES2IMG(A2, "", "rot90")` |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🔍 Built-In Diagnostics & Chemical Engine Log

If a chemical formula returns an error or you wish to verify add-in engine health, click **[SMILES] ➔ [Settings & Diagnostics]**:

* **EPAM Indigo Engine Status:** Confirms native C++ chemistry engine initialization and version.
* **WebView2 & Ketcher Cache Validation:** Verifies SHA256 hashes of embedded offline Ketcher editor assets (69 cached local files).
* **Instant Error Explanation:** When a formula encounters an invalid chemical string (e.g., unclosed rings or impossible valencies), the log displays the exact Indigo parser error message alongside the `#VALUE!` cell indicator.
* **Performance Timing:** Displays rasterization benchmark timings (typically 1.5–3.0 ms per molecule).
