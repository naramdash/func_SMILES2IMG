# Excel Part 1: Structure Tables & Name/CAS Lookup

Turn chemical names, CAS numbers, or SMILES strings into crisp in-cell structure tables across hundreds of Excel rows—and double-click any cell to sketch or edit in place.

<p align="center">
  <img src="/excel-showcase.png" alt="func_SMILES2IMG Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 Download Sample Excel Workbook (.xlsx)</div>
    <p class="s2i-card-desc">Open a ready-made showcase sheet pre-filled with live structure & property formulas</p>
  </a>
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Next: Excel Part 2 (MW, Lipinski, SAR & .sdf) ➔</div>
    <p class="s2i-card-desc">Calculate Molecular Weight, Lipinski Ro5, substructure matches, and import .sdf files</p>
  </a>
</div>

## ⚡ Copy & Paste These 3 Starter Formulas Right Now

Paste any formula below into an empty Excel cell (on Excel 2016–2021, append **`.FLOAT`**, e.g. `=NAME2IMG.FLOAT("Aspirin")`):

```text
=NAME2IMG("Aspirin")
```
```text
=SMILES2IMG(A2)
```
```text
=SMILES2IMG(A2, "", "bw")
```

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

* **Only have a chemical name, lab abbreviation (`THF`, `DMSO`, `DCM`), or CAS number?**  
  Use `=NAME2IMG("Aspirin")`, `=NAME2IMG("THF")`, or `=NAME2IMG("50-78-2")` to render the image directly, or `=NAME2SMILES("Ibuprofen")` to get its canonical SMILES string.
* **Want larger structure images?**  
  Simply increase the Excel **row height (e.g., 80–100 pt)** and **column width**—the molecule scales up automatically while keeping its aspect ratio.

::: details ⚖️ Build an Automated Reaction Stoichiometry & Weighing Table (ELN Template)
Because `=SMILES2MW(B2)` returns a **true numeric value (`double`)**, you can multiply it directly in Excel formulas to calculate **moles (`mmol`) and exact reagent mass (`mg`)** for organic synthesis lab notebooks:

| Col | Header | Row 2 (Limiting Reagent: Salicylic Acid) | Row 3 (Reagent: Acetic Anhydride) | Notes |
| :---: | :--- | :--- | :--- | :--- |
| **Col A** | **Reagent Name** | `Salicylic Acid` | `Acetic Anhydride` | Name or CAS number |
| **Col B** | **SMILES** | `=NAME2SMILES(A2)` | `=NAME2SMILES(A3)` | Auto-resolves SMILES |
| **Col C** | **Structure** | `=SMILES2IMG(B2)` | `=SMILES2IMG(B3)` | Auto-rendered 2D structure |
| **Col D** | **MW (g/mol)** | `=SMILES2MW(B2)` ➔ `138.12` | `=SMILES2MW(B3)` ➔ `102.09` | Numeric value |
| **Col E** | **Equiv (Eq)** | `1.0` | `1.5` | Target molar equivalents |
| **Col F** | **Moles (mmol)** | `5.0` *(Base scale)* | `=$F$2 * E3` ➔ `7.5` | Base mmol × Equiv |
| **Col G** | **Mass to Weigh (mg)** | `=D2 * F2` ➔ **`690.6 mg`** | `=D3 * F3` ➔ **`765.7 mg`** | **MW (Col D) × mmol (Col F)** |
:::

::: details 📚 Built-In 0ms Offline Dictionary (Lab Solvents, Drugs, Biomolecules & CAS Numbers)
The following common laboratory abbreviations, pharmaceuticals, biomolecules, and their CAS Registry Numbers are built directly into the add-in and resolve **instantly in 0ms without internet access**. Any other compound name or CAS number is automatically queried from PubChem once and cached permanently on your PC for offline use.

| Category | Built-In Names, Abbreviations & CAS Numbers (0ms Offline) |
| :--- | :--- |
| **Lab Solvents & Abbreviations** | `THF` (`Tetrahydrofuran`, `109-99-9`), `DMSO` (`Dimethyl Sulfoxide`, `67-68-5`), `DMF` (`Dimethylformamide`, `68-12-2`), `DCM` (`Dichloromethane`, `75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone` (`67-64-1`), `Ethanol` (`64-17-5`), `Methanol` (`67-56-1`), `Isopropanol` (`67-63-0`), `Diethyl Ether` (`60-29-7`), `Chloroform` (`67-66-3`), `Hexane` (`110-54-3`), `Cyclohexane`, `Benzene`, `Toluene`, `Phenol`, `Aniline`, `Pyridine`, `Acetic Acid`, `Formic Acid`, `Benzoic Acid`, `Water` |
| **Pharmaceuticals & Natural Products** | `Aspirin` (`Acetylsalicylic acid`, `50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `Tylenol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen` (`22204-53-1`), `Penicillin G` (`61-33-6`), `Amoxicillin` (`26787-78-0`), `Artemisinin` (`63968-64-9`), `Paclitaxel` (`Taxol`, `33069-62-4`), `Salicylic Acid` (`69-72-7`), `Cholesterol` (`57-88-5`) |
| **Biomolecules, Vitamins & Amino Acids** | `Glucose` (`D-Glucose`, `50-99-7`), `Vitamin C` (`Ascorbic Acid`, `50-81-7`), `Citric Acid` (`77-92-9`), `Lactic Acid` (`50-21-5`), `Glycine` (`56-40-6`), `Alanine` (`56-41-7`), `Dopamine` (`51-61-6`), `Serotonin` (`50-67-9`), `Epinephrine` (`Adrenaline`, `51-43-4`) |
:::

::: info 💡 Which Image Formula Fits My Excel Version?
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(A2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`  
  *(Embeds a true native in-cell picture that sorts and filters seamlessly with your data. Bonus: `=SMILES2IMG(A2)` automatically detects SMILES, InChI, or KET data in `A2`!)*
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** to the formula name!  
  Example: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)` *(Places a cell-locked picture shape that moves and resizes automatically with the cell)*
:::

---

## 2. Common Styling Recipes (B&W · Highlighting · Dark Mode)

Pass optional keywords in the 2nd (`[bg]`) and 3rd (`[style]`) arguments to customize how structures look:

| Desired Style | Ready-to-Copy Formula (`A2`) | Description |
| :--- | :--- | :--- |
| **🎨 Standard Color (Transparent)** | `=SMILES2IMG(A2)` | Standard CPK heteroatom coloring (O = red, N = blue) |
| **📄 ACS Publication B&W** | `=SMILES2IMG(A2, "", "bw")` | Pure black monochrome line art for journal manuscripts |
| **🔍 Substructure Highlight** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | Highlights matching substructure (e.g., benzene ring) in red |
| **🌙 Dark Mode (White Ink)** | `=SMILES2IMG(A2, "#1a1a1a", "white")` | Crisp white bonds on a dark cell background |
| **🔢 NMR Atom Numbering** | `=SMILES2IMG(A2, "", "num stereo")` | Displays atom index numbers (`1, 2, 3...`) and `R/S` stereocenters |

> 🛠️ **Want to build formulas interactively by clicking options?**  
> 👉 Open the [**Interactive Excel Formula Generator & Style Simulator ➔**](/tools/formula-builder) to preview combinations live and copy the formula in one click!

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 3. Draw by Hand & Reaction Schemes (`Ctrl + Shift + K`)

When you want to sketch a structure from scratch or draw a multi-step organic reaction scheme, use the built-in **Ketcher 2D Editor**.

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. Select a cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES] ➔ [Insert / Edit (SMILES)]`** on the ribbon).
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

* **✏️ `Insert / Edit (SMILES)` (`Ctrl+Shift+K`):** Open the 2D molecule sketcher (or choose the adjacent `InChI` / `KET` buttons to specify the formula format)
* **🧬 `3D Structure (Mol*)`:** Open the 3D protein viewer (if the active cell contains a 4-character PDB ID like `1CRN`, it loads that protein automatically!)
* **📋 `Paste Structure`:** Convert clipboard SMILES (or a selected cell containing SMILES text) into a rendered structure formula in 0.1s—and automatically expand row height & column width
* **🖼️ `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** Copy chemical identifiers or a lossless SVG vector for Illustrator / ChemDraw
* **🔬 `Molecule Info`:** Inspect and copy Substance Name, Formula, MW, LogP, Lipinski Ro5, or formatted TSV/Markdown tables
* **🎨 `Batch Style`:** Convert all structures in the sheet to Color, Publication B&W, White Background, or Dark Mode
* **📂 `SDF File`:** Import `.sdf` / `.mol` files into an 8-column Excel table or export the active sheet back to `.sdf`

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: details 🔍 Why does my formula show #VALUE!? (Using Diagnostics)
A `#VALUE!` result usually means the input cell contains a typo in its SMILES string (such as an unclosed ring number or invalid atomic valence).  
Click **`[SMILES] ➔ [Diagnostics]`** on the ribbon to see the exact chemical syntax error message in real time.
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
