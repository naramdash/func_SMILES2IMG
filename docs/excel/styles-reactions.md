# Styles, Highlighting & Reactions

Customize structure rendering per cell using the 2nd (`[bg]`) and 3rd (`[style]`) arguments, or sketch multi-step reaction schemes directly in cells with `=KET2IMG`.

## 1. Formula Styling, Highlighting & Atom Numbering

| Desired Style | Formula Syntax | Use Case |
| :--- | :--- | :--- |
| **Standard Color (Transparent)** | `=SMILES2IMG(A2)` | Standard CPK colors (`O` red, `N` blue) |
| **ACS Publication B&W** | `=SMILES2IMG(A2, "", "bw")` | Pure black line art for journals & patents |
| **Substructure Highlight** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | Highlight matching SMARTS core in color |
| **Dark Mode (White Ink)** | `=SMILES2IMG(A2, "#1E293B", "white")` | Crisp white bonds over dark cell fills |
| **Custom Ink Color** | `=SMILES2IMG(A2, "", "ink=#1e40af")` | Match report or slide accent color |
| **NMR Atom # & Stereo** | `=SMILES2IMG(A2, "", "num stereo")` | 1-based atom indices & `R/S`, `E/Z` labels |

<FeatureMedia id="excel-style-matrix" />

Want to preview options visually before copying? Open the [Interactive Formula Builder](/excel/formula-builder).

## 2. Multi-Step Reaction Schemes (`=KET2IMG` + Ketcher)

Draw organic synthesis schemes with reaction arrows, `+` signs, and catalyst annotations directly in Excel without losing editability:

```text
=KET2IMG(A2)
```

<FeatureMedia id="excel-reactions-ket" />

1. Select an empty cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (or click **`[SMILES]` ➔ `[Insert / Edit (KET)]`**).
2. Draw reactants, reaction arrows (`➔`), catalysts, and products in the offline **Ketcher 2D Editor**.
3. Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to insert `=KET2IMG(...)`. **Double-click** the cell anytime later to reopen the exact reaction layout in `<0.05s`.
