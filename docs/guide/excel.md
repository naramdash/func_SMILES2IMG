# Excel Cookbook Part 1: Structure Tables, Name/CAS, Styles & Reactions

> **Excel Guide Style — Task-Oriented Recipe Cookbook:** Every recipe starts with the live result, gives copy-paste formulas for both **Microsoft 365 / 2024** and **Excel 2016–2021 (`.FLOAT`)**, and pairs a real live screenshot with a recorded Excel walkthrough video.

<FeatureMedia id="excel-overview" />

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 Download 8-Tab Master Showcase Excel Workbook (.xlsx)</div>
    <p class="s2i-card-desc">0.Quick Guide · 1.Structures · 2.Properties/Lipinski/3D · 3.Substructure · 4.Styles Matrix · 5.Reactions · 6.Legacy .FLOAT · 7.Ribbon Tools</p>
  </a>
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Next: Recipes 6–10 (MW, Lipinski, SAR &amp; .SDF) ➔</div>
    <p class="s2i-card-desc">Calculate physicochemical descriptors, screen scaffolds with SMARTS, and import/export .sdf files</p>
  </a>
</div>

## Pick Your Excel Version & 3 Ways to Insert

| Entry Method | How It Works | Best For |
| :--- | :--- | :--- |
| **1. Cell Formulas** | `=SMILES2IMG(A2)`, `=NAME2IMG("Aspirin")`, `=KET2IMG(A2)` | Batch rendering hundreds of rows that sort and filter cleanly |
| **2. Top `[SMILES]` Ribbon** | Click **`[Insert / Edit (SMILES)]`** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>) or **`[Paste Structure]`** | Drawing new molecules/reactions or pasting clipboard SMILES |
| **3. Mouse Actions** | **Double-click** any structure cell to edit in place; **Right-click** ➔ `Chemical Structure` | Rapid edits, copy SVG/SMILES, or opening Molecule Info |

::: info 💡 Microsoft 365 / 2024+ vs. Excel 2016–2021 (`.FLOAT`)
* **Microsoft 365 / Excel 2024+:** Use `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)` for native **in-cell images** that sort and filter with table rows.
* **Excel 2016 / 2019 / 2021:** Append **`.FLOAT`** (`=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)`) to create cell-anchored floating shapes (`xlMoveAndSize`).
:::

---

## Recipe 1: Turn a SMILES Column into a Scaled Structure Table

**Goal:** Keep compound names in Column A and SMILES strings in Column B, then render crisp, aspect-ratio-locked 2D structures in Column C alongside formula and molecular weight columns.

```text
=SMILES2IMG(B2)
```

<FeatureMedia id="excel-formula-table" />

1. Put compound names in `A2:A5` and SMILES strings in `B2:B5` (e.g., `CC(=O)Oc1ccccc1C(=O)O` for Aspirin).
2. Type `=SMILES2IMG(B2)` in `C2` and press <kbd>Enter</kbd>.
3. Increase **row height (e.g. `62–80 pt`)** and **column width**—molecules scale automatically without distortion.
4. Drag the fill handle down (`C2:C5`) to render hundreds of rows in milliseconds.

---

## Recipe 2: Render Directly from Chemical Names, Lab Solvents & CAS Numbers

**Goal:** Skip looking up SMILES codes by passing common compound names, lab abbreviations (`THF`, `DMSO`, `DCM`), or CAS numbers directly to `=NAME2IMG` or `=NAME2SMILES`.

```text
=NAME2IMG("Aspirin")
=NAME2IMG("50-78-2")
=NAME2SMILES("Ibuprofen")
```

<FeatureMedia id="excel-name-cas" />

::: details 📚 Built-In 0ms Offline Dictionary (Solvents, Drugs, Biomolecules & CAS Numbers)
Common lab solvents, pharmaceuticals, biomolecules, and CAS Registry Numbers resolve **in 0ms completely offline**. Any uncached compound is looked up once from PubChem and cached locally on disk for future offline use:
* **Lab Solvents:** `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`
* **Pharmaceuticals:** `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel`
* **Biomolecules:** `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`
:::

::: details ⚖️ Bonus ELN Stoichiometry Table Using `=SMILES2MW` (Numeric Double)
Because `=SMILES2MW` returns a true Excel number (`double`), multiply it directly by `mmol` to compute exact reagent weights (`mg`):
| Col A (Name) | Col B (`=NAME2SMILES(A2)`) | Col C (`=SMILES2IMG(B2)`) | Col D (`=SMILES2MW(B2)`) | Col E (Equiv) | Col F (`mmol`) | Col G (`=D2*F2` mg) |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D structure)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D structure)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |
:::

---

## Recipe 3: ACS Monochrome (B&W), Substructure Highlighting, Dark Mode & Atom Numbering

**Goal:** Format structures for patents, print journals, NMR peak assignments, or dark-themed dashboards using the 2nd (`[bg]`) and 3rd (`[style]`) formula arguments.

| Desired Style | Formula Syntax | Use Case |
| :--- | :--- | :--- |
| **Standard Color (Transparent)** | `=SMILES2IMG(A2)` | Standard CPK colors (`O` red, `N` blue) |
| **ACS Publication B&W** | `=SMILES2IMG(A2, "", "bw")` | Pure black line art for journals & patents |
| **Substructure Highlight** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | Highlight matching SMARTS core in color |
| **Dark Mode (White Ink)** | `=SMILES2IMG(A2, "#1E293B", "white")` | Crisp white bonds over dark cell fills |
| **NMR Atom # & Stereo** | `=SMILES2IMG(A2, "", "num stereo")` | 1-based atom indices & `R/S`, `E/Z` labels |

<FeatureMedia id="excel-style-matrix" />

> 🛠️ Want to click options visually? Open the [**Interactive Formula Builder**](/tools/formula-builder).

---

## Recipe 4: Sketch & Persist Multi-Step Reaction Schemes (`=KET2IMG` + Ketcher)

**Goal:** Draw organic synthesis schemes with reaction arrows, `+` signs, and catalyst annotations directly in Excel without losing editability.

```text
=KET2IMG(A2)
```

<FeatureMedia id="excel-reactions-ket" />

1. Select an empty cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES] ➔ [Insert / Edit (KET)]`**).
2. Draw reactants, reaction arrows (`➔`), catalysts, and products in the offline **Ketcher 2D Editor**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert `=KET2IMG(...)`. **Double-click** the cell later to reopen the exact reaction coordinates in `<0.05s`.

---

## Recipe 5: Excel 2016–2021 Compatibility Mode (`.FLOAT`)

**Goal:** Render cell-anchored chemical images on Excel 2016, 2019, or 2021 (where native in-cell picture formulas are unavailable).

```text
=SMILES2IMG.FLOAT(A2)
=NAME2IMG.FLOAT("Caffeine")
```

<FeatureMedia id="excel-legacy-float" />

Floating picture shapes automatically track their host cell (`xlMoveAndSize`), so resizing row height or column width scales the structure cleanly.

---

## Excel `[SMILES]` Ribbon & Right-Click Tools

<FeatureMedia id="excel-ribbon-tools" />

* **`Insert / Edit (SMILES / InChI / KET)` (`Ctrl+Shift+K`):** Launch Ketcher for the selected cell and emit the matching formula format.
* **`3D Structure (Mol*)`:** Launch the 3D viewer (auto-loads a 4-letter PDB code like `1CRN` if present in the active cell).
* **`Paste Structure`:** Convert clipboard or selected-cell SMILES/InChI/KET into a structure formula in 0.1s and auto-size the cell.
* **`Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** Extract identifiers or lossless vector SVG directly from any chemical cell.
* **`Molecule Info`:** Open the 17-property dialog or copy TSV/Markdown/JSON rows.
* **`Batch Style`:** Restyle your selected range (or entire active sheet) to `Standard Color`, `Publication B&W`, `White Background`, or `Dark Mode`.
* **`SDF File`:** Import multi-compound `.sdf` libraries into a 12-column table or export sheets back to `.sdf`.
* **`Diagnostics` & `About func_SMILES2IMG`:** Inspect live parser logs (`#VALUE!` causes) or check for updates.

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Live 3-Level Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🚀 Continue in Excel

<div class="s2i-card-grid">
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel Recipes 6–10: Cheminformatics &amp; SDF</div>
    <p class="s2i-card-desc">MW, Exact Mass, Lipinski Ro5, 12-group SMARTS dictionary, similarity &amp; .sdf I/O</p>
  </a>
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📑 Complete 40-Formula Cheat Sheet</div>
    <p class="s2i-card-desc">Every signature, return data type, alias token, and error code reference</p>
  </a>
</div>
