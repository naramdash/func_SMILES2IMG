# Excel Basics: Structure Tables & Styling

Learn how to render hundreds of chemical structures directly inside Excel cells, customize styling (ACS B&W, substructure highlights, dark mode), and edit molecules in place by double-clicking.

> 🔬 **Looking for Molecular Weight, Exact Mass, Lipinski Ro5, SAR Substructure Screening, or `.sdf` File Import/Export?**  
> 👉 Jump straight to **[Excel Cheminformatics: Properties, SAR Screening & SDF Files ➔](/guide/excel-cheminformatics)**!

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

| Column | Content | Formula in Row 2 | Example Output (`A2` = Aspirin) |
| :---: | :--- | :--- | :--- |
| **Col A** | Input (SMILES or Name) | `CC(=O)Oc1ccccc1C(=O)O` | Input text |
| **Col B** | **2D Structure Image** | `=SMILES2IMG(A2)` | *(Auto-scaled in-cell structure)* |
| **Col C** | **Chemical Formula** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **Col D** | **Molecular Weight (g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **Col E** | **Lipinski Rule of 5** | `=SMILES2LIPINSKI(A2)` | `PASS` |

* **Only have a chemical name or CAS number?**  
  Use `=NAME2IMG("Aspirin")` or `=NAME2IMG("50-78-2")` to render the image directly, or `=NAME2SMILES("Ibuprofen")` to get its canonical SMILES string (powered by a 60+ reagent offline dictionary + PubChem lookup).
* **Want larger structure images?**  
  Simply increase the Excel **row height (e.g., 80–100 pt)** and **column width**—the molecule scales up automatically while keeping its aspect ratio.

::: info 💡 Which Image Formula Fits My Excel Version?
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(A2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`  
  *(Embeds a true native in-cell picture that sorts and filters seamlessly with your data)*
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** to the formula name!  
  Example: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)` *(Places a cell-locked picture shape that moves and resizes automatically with the cell)*
:::

---

## 2. Interactive Styling Options (B&W · Highlighting · Dark Mode)

Click any option below to preview how it looks and copy the exact Excel formula in one click!

<FormulaBuilder />

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 3. Draw by Hand & Reaction Schemes (`Ctrl + Shift + K`)

When you want to sketch a structure from scratch or draw a multi-step organic reaction scheme, use the built-in **Ketcher 2D Editor**.

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. Select a cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES] ➔ [Insert Structure]`** on the ribbon).
2. Draw rings, functional groups, stereo bonds (wedges/hashes), or **reaction arrows and catalyst text (H₂SO₄, 55 °C)**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert into the cell immediately:
   * Single molecules are inserted as `=SMILES2IMG(...)`.
   * Multi-step reactions with arrows or text are saved losslessly as `=KET2IMG(...)`.
4. **Need to make changes later?** Simply **double-click** the cell to reopen the exact molecule or reaction scheme in under 0.05 seconds.

---

## 4. Top `[SMILES]` Ribbon & Right-Click Menu Summary

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ Insert Structure (`Ctrl+Shift+K`):** Open the 2D molecule & reaction sketcher
* **🧬 3D Structure (Mol\*):** Open the 3D protein viewer (if the active cell contains a 4-character PDB ID like `1CRN`, it loads that protein automatically!)
* **⚡ Quick Insert:** Convert a cell containing SMILES text into a rendered structure formula in 0.1s
* **🖼️ Copy SVG:** Copy the selected molecule as a lossless SVG vector for Illustrator or ChemDraw
* **🎨 Batch Style:** Convert all structures in the sheet to Color, Publication B&W, White Background, or Dark Mode
* **📂 SDF File:** Import multi-molecule `.sdf` files into formatted Excel tables or export back to `.sdf`

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: details 🔍 Why does my formula show #VALUE!? (Using Diagnostics)
A `#VALUE!` result usually means the input cell contains a typo in its SMILES string (such as an unclosed ring number or invalid atomic valence).  
Click **`[SMILES] ➔ [Settings & Diagnostics]`** on the ribbon to see the exact chemical syntax error message in real time.
:::

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel Cheminformatics: SAR & SDF</div>
    <p class="s2i-card-desc">MW, Exact Mass (HRMS), Lipinski Ro5, 12-group SMARTS dictionary, similarity ranking & .sdf files</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 Publishing, SVG & Property Tables</div>
    <p class="s2i-card-desc">1-click ACS B&W sheet conversion, lossless Illustrator SVG copy, and 5 property export formats</p>
  </a>
</div>
