# Ribbon & Context Menu

Reference for every button on the Excel **`[SMILES]`** ribbon tab and the 3-level **`Chemical Structure`** right-click menu.

<FeatureMedia id="excel-ribbon-tools" />

## Top `[SMILES]` Ribbon Commands

| Group | Command | Action |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit (SMILES / InChI / KET)` (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>) | Launch Ketcher for the active cell and write the matching formula |
| | `3D Structure (Mol*)` | Launch Mol\* 3D viewer (auto-loads a 4-letter PDB code if present in the active cell) |
| | `Paste Structure` | Convert clipboard or selected-cell SMILES/InChI/KET into a structure formula in 0.1s |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG` | Copy identifiers or lossless vector SVG directly from any chemical cell |
| | `Molecule Info` | Open the 17-property window or copy TSV, Summary, Multi-Line, Markdown, or JSON |
| | `Batch Style` | Restyle selection or entire sheet (`Standard Color`, `Publication B&W`, `White Background`, `Dark Mode`) |
| | `SDF File` | Import 12-column `.sdf` tables or export the active sheet to V2000 `.sdf` |
| **Diagnostics & About** | `Diagnostics` | Inspect live EPAM Indigo parser error logs (`#VALUE!` causes) |
| | `About func_SMILES2IMG` | View version, third-party licenses, and check for updates |

## Right-Click Context Menu

![Live 3-Level Chemical Structure Context Menu in Excel](/excel-context-menu-guide.png)
