# Microsoft Excel Guide

Learn how to render hundreds of chemical structures directly inside Excel cells, automatically calculate molecular weight and drug-likeness (Lipinski Ro5), highlight substructures, and import/export `.sdf` files.

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ Excel in 4 Simple Steps
1. **Render by Formula:** Type `=SMILES2IMG(A2)` or `=NAME2IMG("Aspirin")` and drag down across rows.
2. **Draw by Hand:** Select any cell, press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>, sketch your molecule or reaction, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
3. **Double-Click to Edit:** **Double-click** any molecule cell to reopen the sketcher and update it in place.
4. **Right-Click to Copy Info:** **Right-click ➔ `Chemical Structure`** to copy MW, Formula, or lossless SVG vectors in one click.
:::

---

## 1. Build a Compound & Property Table with Formulas

The fastest way to build a chemical table in Excel is to put SMILES strings (or chemical names / CAS numbers) in Column A and enter formulas in the adjacent columns.

### Example Worksheet Setup

| Column | Content | Formula in Row 2 | Example Output (A2 = `"Aspirin"` or SMILES) |
| :---: | :--- | :--- | :--- |
| **Col A** | Input (SMILES or Name) | `CC(=O)Oc1ccccc1C(=O)O` | Input text |
| **Col B** | **2D Structure Image** | `=SMILES2IMG(A2)` | *(Auto-scaled in-cell structure)* |
| **Col C** | **Chemical Formula** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **Col D** | **Molecular Weight (g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **Col E** | **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `PASS` |

* **Only have a chemical name or CAS number?**  
  Use `=NAME2IMG("Aspirin")` or `=NAME2IMG("50-78-2")` to render the image directly, or `=NAME2SMILES("Ibuprofen")` to get its canonical SMILES string.
* **Want larger structure images?**  
  Simply increase the Excel **row height (e.g., 80–100 pt)** and **column width**—the molecule scales up automatically while keeping its aspect ratio.

::: info 💡 Which Image Formula Fits My Excel Version?
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(A2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`  
  *(Embeds a true native in-cell picture that sorts and filters seamlessly with your data)*
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** to the formula name!  
  Example: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)` *(Places a cell-locked picture shape that moves and resizes automatically with the cell)*
:::

---

## 2. Draw by Hand & Reaction Schemes (`Ctrl + Shift + K`)

When you want to sketch a structure from scratch or draw a multi-step organic reaction scheme, use the built-in **Ketcher 2D Editor**.

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. Select a cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES] ➔ [Insert Structure]`** on the ribbon).
2. Draw rings, functional groups, stereo bonds (wedges/hashes), or **reaction arrows and catalyst text ($H_2SO_4$, 55 °C)**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert into the cell immediately:
   * Single molecules are inserted as `=SMILES2IMG(...)`.
   * Multi-step reactions with arrows or text are saved losslessly as `=KET2IMG(...)`.
4. **Need to make changes later?** Simply **double-click** the cell to reopen the exact molecule or reaction scheme in Ketcher.

---

## 3. Substructure Highlighting & SAR Screening

Combine zero-latency chemical formulas with Excel's **AutoFilter** and **Conditional Formatting** to screen compound libraries directly in your spreadsheet.

### ① Highlight a Core Scaffold or Functional Group (`hl:`)
Pass `"hl:<substructure>"` in the 3rd argument to highlight matching atoms and bonds:
```text
=SMILES2IMG(A2, "", "hl:c1ccccc1")
```

### ② 3-Step Virtual Screening Recipe
1. **Filter by Substructure (`TRUE` / `FALSE`):**
   ```text
   =SMILES.MATCH(A2, "c1ccccc1")
   ```
2. **Rank by 2D Similarity against a Lead Compound in `$B$1` (`0.0 ~ 1.0`):**
   ```text
   =SMILES.SIMILARITY(A2, $B$1)
   ```
3. **Evaluate Drug-Likeness (Lipinski Ro5) & Physicochemical Properties:**
   * Pass/Fail check (`"PASS"` / `"FAIL"`): `=SMILES2LIPINSKI(A2)`
   * LogP & Polar Surface Area (TPSA): `=SMILES2LOGP(A2)`, `=SMILES2TPSA(A2)`
   * Monoisotopic Exact Mass (for HRMS/MS): `=SMILES2EXACTMASS(A2)`

> 📐 See all 25+ worksheet functions in the **[Excel Formula Cheat Sheet](/reference/formula)**.

---

## 4. Popular Styling Options (B&W · Dark Mode · Atom Numbers)

Combine any option keywords in the 3rd argument of `=SMILES2IMG(cell, [background], [options])` using spaces or commas:

| Desired Style | Example Formula | What It Does |
| :--- | :--- | :--- |
| **ACS / Patent Black & White** | `=SMILES2IMG(A2, "", "bw")` | High-contrast monochrome line art for journal manuscripts and patents |
| **Substructure Highlight** | `=SMILES2IMG(A2, "", "hl:C(=O)O")` | Highlights specific functional groups or scaffolds in color |
| **Atom Numbers & Stereo Labels** | `=SMILES2IMG(A2, "", "num stereo")` | Shows atom indices (`num`, great for NMR) and *R/S*, *E/Z* labels (`stereo`) |
| **Dark Mode (White Ink)** | `=SMILES2IMG(A2, "#1e1e1e", "white")` | Crisp white bonds rendered over a dark cell background |
| **Explicit Hydrogens / Rotate / Flip** | `=SMILES2IMG(A2, "", "h rot90 flip")` | Unfolds hydrogens (`h`), rotates 90° (`rot90`), or flips horizontally (`flip`) |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 5. Top `[SMILES]` Ribbon Tab & `.sdf` File Exchange

Use the **`[SMILES]`** tab on the Excel ribbon for one-click table imports, whole-sheet restyling, and vector export:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### 📂 Import & Export `.sdf` Chemical Files (`[SDF File]`)
* **📥 Import `.sdf` File...:** Select any multi-molecule `.sdf` file to automatically generate a formatted Excel table containing **rendered structure images (`=SMILES2IMG`), SMILES, Formula, Molecular Weight, and all SD property tags (`<IC50>`, `<Activity>`, etc.)**.
* **📤 Export `.sdf` File...:** Exports your worksheet compound table back to a standard V2000 `.sdf` file.

### 🎨 Restyle the Entire Sheet in 1 Click (`[Batch Style]`)
Unify every chemical structure in your selection or active sheet to match your target document:
* **Standard Color:** IUPAC element colors with a transparent background (default for spreadsheets)
* **Publication B&W:** High-contrast ACS / patent monochrome line art
* **White Background:** Element colors on a solid white background (stays clear over colored or banded rows)
* **Dark Mode:** Crisp white bonds on a dark/transparent background

### 🖼️ Copy Vector Graphics for Illustrator (`[Copy SVG]`)
* Copies the selected molecule as a lossless **scalable vector graphic (`SVG`)** to the clipboard, ready to paste directly into Adobe Illustrator, Inkscape, or ChemDraw.

---

## 6. Right-Click Context Menu (`Chemical Structure`)

Right-click any cell to access quick chemical actions under the **`Chemical Structure`** menu:

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ Edit Structure in Ketcher:** Reopen the molecule in the 2D sketcher for in-place editing
* **🧬 3D Structure (Mol\*)...:** Open the 3D protein viewer (if the active cell contains a 4-character PDB ID like `1CRN`, it loads that protein automatically!)
* **📋 Copy SMILES / InChI / KET / SVG:** Copy chemical notations or vector graphics in one click
* **🔬 Copy Molecule Info (Submenu & Summary Window):**
  * Copy Substance Name, Formula, Molecular Weight, or InChIKey individually,
  * Copy an entire **Excel Row (TSV)**, **Markdown Table**, or **Bulleted List** to paste across adjacent cells or documents,
  * Or open the **Detailed Molecule Info Window** to inspect all properties (including LogP, TPSA, and Lipinski Ro5) and export as JSON.

::: details 🔍 Why does my formula show #VALUE!? (Using Diagnostics)
A `#VALUE!` result usually means the input cell contains a typo in its SMILES string (such as an unclosed ring number or invalid atomic valence).  
Click **`[SMILES] ➔ [Settings & Diagnostics]`** on the ribbon to see the exact chemical syntax error message in real time.
:::
