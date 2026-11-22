# Styles & Substructure Highlights

All 8 structure image functions (**`=SMILES2IMG`**, **`=NAME2IMG`**, **`=INCHI2IMG`**, **`=KET2IMG`**, and their `.FLOAT` equivalents) use the exact same 2nd (`[background]`), 3rd (`[style]`), and 4th (`[transform]`) arguments—so one style vocabulary controls monochrome journals, dark mode, highlights, and labels across every formula.

<FormulaTable
  title="Styling & Highlighting Sheet"
  :columns='["SMILES", "Formula"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A2, \"\", \"bw\")", desc: "ACS monochrome black-and-white line art", out: "B&W image" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A3, \"#1E293B\", \"white\")", desc: "Dark slate cell background with white bonds", out: "Dark Mode" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A4, \"\", \"hl:c1ccccc1\")", desc: "Color highlight on benzene substructure atoms", out: "Highlighted" }
  ]'
/>

## 1. Shared 4-Argument Anatomy

| Argument | Position | Default | Supported Values |
| :--- | :---: | :--- | :--- |
| `input` | 1st | *(Required)* | SMILES, Name/CAS, InChI, or KET cell/string |
| `[background]` | 2nd | `""` (`trans`) | `""` / `"trans"`, `"white"`, CSS color (`"yellow"`), or `#RRGGBB` hex fill |
| `[style]` | 3rd | `""` (`color`) | `"color"`, `"bw"`, `"white"`, `"ink=#HEX"`, `"hl:<smarts>"`, `"num"`, `"stereo"`, `"h"`, `"all-c"` |
| `[transform]` | 4th | `""` | `"flip"`/`"flipx"`, `"flipy"`, `"rot90"`, `"rot180"`, `"rot270"` *(also accepted inside `[style]`)* |

::: tip Shortcut for Monochrome (`"bw"`)
Pass `"bw"` (or `FALSE`) directly in the 2nd argument to skip empty quotes: `=SMILES2IMG(A2, "bw")` and `=NAME2IMG("Aspirin", "bw")`.
:::

## 2. Copy-Ready Style, Highlight & Transform Matrix

<FeatureMedia id="excel-style-matrix" />

| Target Look | Formula Example | Visual Effect |
| :--- | :--- | :--- |
| **Standard IUPAC Color** | `=SMILES2IMG(A2)` | Colored heteroatoms (`O` red, `N` blue) over transparent cells |
| **ACS Journal Monochrome** | `=SMILES2IMG(A2, "", "bw")` | Publication black-and-white line art (`mono`, `black`, `grayscale`) |
| **Opaque White Card Box** | `=SMILES2IMG(A2, "white", "bw")` | Solid white fill—blocks shaded cell backgrounds behind the drawing |
| **Dark Mode (White Ink)** | `=SMILES2IMG(A2, "#1E293B", "white")` | Crisp white bonds over dark navy/slate dashboard cells |
| **Custom Brand Ink Color** | `=SMILES2IMG(A2, "", "ink=#1E40AF")` | Custom hex/CSS bond color (`line=#HEX`, `base=#HEX`) |
| **Substructure Highlight** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | Colors matching SMILES/SMARTS substructure atoms and bonds |
| **Dynamic Header Highlight** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | Edit query cell `$B$3` once to re-highlight an entire screening table |
| **NMR Atom Numbers & Stereo** | `=SMILES2IMG(A2, "", "num stereo")` | 1-based atom indices (`1, 2, 3...`) and CIP *R/S*, *E/Z* descriptors |
| **Explicit H & Carbon Vertices** | `=SMILES2IMG(A2, "", "h all-c")` | Unfolds explicit `-H` atoms and draws `C` symbols on carbons |
| **Mirror Flip & Rotation** | `=SMILES2IMG(A2, "", "flip rot90")` | Horizontal/vertical flip and `90°`/`180°`/`270°` clockwise rotation |

## 3. Combine Tokens & Restyle Whole Sheets

* **Flexible Delimiters:** Separate multiple `[style]` tokens using spaces (`"bw num"`), commas (`"bw, num"`), pipes (`"bw|num"`), semicolons (`"bw;num"`), or `+` (`"bw+num"`).
* **1-Click Sheet Conversion:** Restyle hundreds of existing structure formulas at once with **`[SMILES]` ➔ `[Batch Style]`** ([Batch Style, TSV, 3D](/excel/tools-export)).

## Related Guides

- [**All 40 Functions**](/excel/formulas) — Full parameter catalog and accepted style aliases
- [**Substructure & Similarity**](/excel/sar) — Pair `hl:<SMARTS>` with `=SMILES.MATCH` and the 12-group SMARTS dictionary
- [**Batch Style, TSV, 3D**](/excel/tools-export) — Convert an entire sheet cleanly to ACS B&W or Dark Mode
