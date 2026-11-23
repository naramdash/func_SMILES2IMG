# Word Overview

Sketch, convert, style, and re-edit 2D structures, reaction schemes, and 3D protein figures directly inside Microsoft Word—using open-source tools and standard `.docx` media without changing table layouts when substituents change.

<FeatureMedia id="word-overview" />

## Key Workflows

### 1. Double-Click In-Place Edits Without Breaking SAR Table Widths
Fix a methyl group or swap a catalyst inside a multi-row table by **double-clicking** the structure. `func_SMILES2IMG` keeps the aspect ratio locked and updates the image inside the existing cell bounds.

<FeatureMedia id="word-ch3-compound-table" />

### 2. ACS Monochrome Batch Styling & Illustrator Vector `SVG`
Switch colored draft figures across an entire `.docx` (or a highlighted table) to black-and-white line art via **`[Batch Style]`**, or copy vector `SVG` into Adobe Illustrator for TOC graphics.

<FeatureMedia id="word-ch6-publication-finish" />

### 3. Calculated HRMS / Formula Copy & 3D Protein Figures (`Mol*`)
Copy Chemical Formula, Molecular Weight, Monoisotopic Exact Mass (`Da`), LogP, and TPSA straight from any structure into your Experimental section, or enter a 4-character PDB ID (`1CRN`, `6LU7`) to insert a 3D protein figure.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## Comparison: Word Manuscript Authoring

| Manuscript Task | Without `func_SMILES2IMG` | With `func_SMILES2IMG` |
| :--- | :--- | :--- |
| **Editing a structure inside a compound table** | Reopen external editor → export PNG → delete & reinsert → fix shifted column widths | **Double-click the structure → edit in Ketcher 2D → <kbd>Ctrl</kbd> + <kbd>Enter</kbd>** |
| **Switching color figures to ACS B&W** | Re-export and replace every file one by one | **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`** |
| **Entering Formula, MW & HRMS Exact Mass** | Look up or calculate manually in another window | **`[Molecule Info]` copies 12 descriptors & Exact Mass** |
| **Sharing `.docx` with Mac co-authors & reviewers** | Proprietary OLE objects can break or fail to edit on macOS | **Standard high-DPI PNG + `.glb` embedded—readable on Mac, Web & PDF** |
| **Adding a 3D protein binding-site figure** | Launch PyMOL / ChimeraX, render PNG, crop and re-import manually | **`[3D Structure (Mol*)]` fetches PDB by ID & preserves camera angle on re-edit** |

## Architecture & Compatibility

<OfficeOverviewCta app="word" />

## Quick Walkthrough

1. **Text-to-structure conversion:** Type `Aspirin` or `CC(=O)Oc1ccccc1C(=O)O` in a blank Word paragraph, highlight it, and click **`[SMILES]` ➔ `[Paste Structure]`** (or right-click ➔ **`Chemical Structure` ➔ `Paste Structure`**).
2. **Sketch or edit in place:** **Double-click** that molecule to reopen **Ketcher 2D**, add a substituent, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>—or click **`[SMILES]` ➔ `[Insert / Edit Structure]`** on an empty line to draw from scratch.
3. **Practice inside the Playbook (`.docx`):** Download <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Word_Playbook.docx" download="func_SMILES2IMG_Word_Playbook.docx"><strong>func_SMILES2IMG_Word_Playbook.docx</strong></a> and work through 6 guided pages of SAR tables, KET reaction schemes, Mol\* `1CRN` figures, and ACS B&W conversion ([see 6-page breakdown →](/word/playbook)).

## Jump to a Word Workflow

| What You Want to Do in Word | Guide | Key Command |
| :--- | :--- | :--- |
| See manuscript, patent, SI, structural & thesis scenarios | [Word Use Cases](/word/use-cases) | 5 R&D & thesis scenarios |
| Sketch molecules, convert typed SMILES/names, and build lossless `KET` reaction schemes | [1. 2D Structures & Reaction Schemes](/word/structures) | `[Insert / Edit Structure]` / `[Paste Structure]` |
| Lock SAR table widths, fix clipped top edges, and copy Formula / MW / HRMS / SI lists | [2. SAR Tables & Experimental Properties](/word/tables-properties) | Fixed Width + `[Molecule Info]` |
| Fetch PDB IDs or local `.pdb`/`.cif` files as 2D PNGs or interactive 3D `.glb` models | [3. 3D Protein Figures (Mol*)](/word/3d-figures) | `[3D Structure (Mol*)]` |
| Run the pre-submission checklist (ACS B&W, TOC SVG, SI extraction, Mac sharing) | [4. Journal Finish, SVG, Sharing](/word/publishing) | `[Batch Style]` / `[Copy SVG]` |
| Inspect every `[SMILES]` ribbon button and the 5 right-click context menus | [Ribbon & Context Menus](/word/reference) | Top `[SMILES]` tab & Right-Click |
