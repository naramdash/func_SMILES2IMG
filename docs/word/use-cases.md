# Word Use Cases

Explore 5 real-world ways **`func_SMILES2IMG`** streamlines Microsoft Word workflows—from journal submission and fixed-width patent SAR tables to Supporting Information (SI) generation, 3D binding-pocket figures, and 100+ page theses.

| Use Case | What You Can Do | Key Tool |
| :--- | :--- | :--- |
| [**1. Journal Manuscript & ACS B&W**](#journal-manuscript) | Convert a color draft to monochrome line art and export vector TOC graphics | `[Batch Style]`, `[Copy SVG]` |
| [**2. Patent & SAR Compound Tables**](#patent-sar) | Keep table widths locked during structure conversion and calculated property fill | `[Paste Structure]`, `[Molecule Info]` |
| [**3. Synthesis Reports & SI Blocks**](#lab-report-si) | Embed lossless editable `KET` schemes and paste characterization lists | `[Insert / Edit (KET)]`, `Multi-Line List` |
| [**4. 3D Protein & Pocket Reports**](#structural-reports) | Insert PDB/local `.pdb` figures (`PNG` or 3D `.glb`) and re-edit camera angles anytime | `[3D Structure (Mol*)]` |
| [**5. Long Thesis & Dual Print/PDF**](#thesis-dissertation) | Re-edit or batch-restyle 100+ page documents without breaking Word Captions or TOF links | Double-Click + `[Batch Style]` |

---

## 1. Journal Manuscript & ACS B&W {#journal-manuscript}

Draft with colored structures while revising with macOS co-authors, then switch your `.docx` to monochrome black-and-white and export a vector Graphical Abstract (TOC) before uploading.

<FeatureMedia id="word-ch6-publication-finish" />

1. **Draft in Default Color:** Insert inline molecules and `KET` reaction schemes using standard IUPAC element colors while iterating with co-authors.
2. **Monochrome Batch Finish:** Place your cursor in prose (nothing selected) and click **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`**. Every 2D figure switches to black line art without disturbing table widths (`Mol*` 3D objects stay colored).
3. **Illustrator Vector TOC Graphic:** Select your key reaction scheme, click **`[SMILES]` ➔ `[Copy SVG]`**, and press <kbd>Ctrl</kbd> + <kbd>V</kbd> in Adobe Illustrator or Inkscape for editable vector paths.
4. **Share Without Broken Links:** Send `.docx` directly to macOS co-authors or editorial portals—embedded high-DPI PNGs render identically everywhere while preserving `Alt Text` tags for future Windows edits.

* **Next:** [**4. Journal Finish, SVG, Sharing**](/word/publishing) · Playbook **Page 6**

---

## 2. Patent & SAR Compound Tables {#patent-sar}

Keep multi-page patent claims and SAR summary tables crisp and aligned even as substituents change—eliminating clipped top bonds, collapsed column widths, and manual transcription typos in Formula, MW, or HRMS values.

<FeatureMedia id="word-ch3-compound-table" />

1. **Lock Table Geometry:** Right-click the table ➔ **AutoFit ➔ Fixed Column Width**, and ensure cell paragraph line spacing is **Single** or **Multiple** (never *Exactly*, which slices tall rings).
2. **Convert Typed Lists In-Place:** Paste chemical names (`Aspirin`), CAS numbers (`50-78-2`), or SMILES strings into column cells, highlight each cell, and click **`[SMILES]` ➔ `[Paste Structure]`**.
3. **Fill Calculated Descriptors:** Run **`[Molecule Info]`** ➔ **`Copy as Excel Row (TSV Columns)`** to populate adjacent cells, or copy **`Copy Chemical Formula`** and Exact Mass (`Da`) individually.
4. **Re-Edit Claims In Place:** **Double-click** any table structure to modify rings or linkers—`func_SMILES2IMG` replaces the image inside the exact locked column width.

* **Next:** [**2. SAR Tables & Experimental Properties**](/word/tables-properties) · Playbook **Page 3**

---

## 3. Synthesis Reports & SI Blocks {#lab-report-si}

Keep every synthetic scheme inside your lab reports and Supporting Information (SI) permanently editable alongside machine-verified Formula, Molecular Weight, HRMS Exact Mass, and SMILES/InChI blocks.

<FeatureMedia id="word-ch4-reaction-scheme" />

1. **Embed Lossless `KET` Schemes:** Click **`[SMILES]` ➔ `[Insert / Edit Structure]`**, draw reactants, arrows (`➔`), `+` signs, and `T`-tool conditions (`H₂SO₄, 85 °C, 92%`), then press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>. Full coordinates stay embedded (`KET:...`) for double-click updates weeks later.
2. **Paste SI Characterization Lists:** Select a compound picture—or highlight a SMILES string in prose—and click **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** to paste Formula, MW, Exact Mass (`Da`), LogP, TPSA, SMILES, InChI, and InChIKey cleanly.
3. **Copy Portal Strings:** Use **`[Copy SMILES]`** or **`[Copy InChI]`** on the ribbon whenever journal portals ask for compound spreadsheet strings.

* **Next:** [**1. 2D Structures & Reaction Schemes**](/word/structures) · Playbook **Pages 2 & 4**

---

## 4. 3D Protein & Pocket Reports {#structural-reports}

Embed publication-quality binding-pocket views from RCSB PDB IDs or local proprietary docking files (`.pdb`, `.cif`), and adjust viewing angles later without reopening external 3D viewers from scratch.

<FeatureMedia id="word-ch5-edit-and-molstar" />

1. **Preload From Highlighted Text:** Highlight a 4-letter PDB code (`1CRN`, `6LU7`) in your draft and click **`[SMILES]` ➔ `[3D Structure (Mol*)]`**, or click **`Open File...`** inside Mol\* to load local `.pdb`/`.cif`/`.sdf` files offline.
2. **Choose Output Format:** Insert **`2D Picture`** (`BG: White` or `Transparent`) for journal PDFs, or **`3D Model (GLB)`** for interactive in-page rotation in internal `.docx` reviews (Word 2019+/365).
3. **Refine Camera Angles Later:** **Double-click** any placed Mol\* figure to reload the saved camera coordinates, colors, and representation from `CustomXMLParts`.

::: tip Cross-App Suite Handoff
Screen candidates in [**Excel Use Cases**](/excel/use-cases) ➔ finalize tables and `KET` schemes in Word ➔ copy with **`[Copy KET]`** / **`[Copy SMILES]`** onto dark slides in [**PowerPoint Use Cases**](/powerpoint/use-cases).
:::

* **Next:** [**3. 3D Protein Figures (`Mol*`)**](/word/3d-figures) · Playbook **Page 5**

---

## 5. Long Thesis & Dual Print/PDF {#thesis-dissertation}

Manage 100+ page theses safely when every scheme is linked to Word **Captions** (`Scheme 3.2`), cross-references, and a Table of Figures—and maintain both a color digital defense PDF and a black-and-white printed bound edition.

<FeatureMedia id="word-usecase-thesis-captions" />

1. **Preserve Captions & Cross-References:** Attach standard Word **Insert Caption** numbering beneath each inline figure. Both **double-click** revisions and **`[Batch Style]`** update the picture in place at its exact paragraph anchor without touching caption numbers or TOF links.
2. **Produce Color PDF & B&W Bound Copies:** Keep one color master `.docx` for defense review; save a duplicate for library printing and run **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** across the thesis.
3. **No External `.cdx` Source Needed Months Later:** Because `SMILES` and `KET` metadata stay embedded inside the `.docx` itself, post-defense edits months later require only a double-click.

* **Next:** [**1. 2D Structures & Reaction Schemes**](/word/structures) · [**4. Journal Finish, SVG, Sharing**](/word/publishing)

---

## Next Steps

- [**Practice in the 6-Page Word Playbook (.docx)**](/word/playbook) — Try `Paste Structure`, fixed SAR tables, KET reactions, `1CRN` 3D views, and ACS B&W in Word.
- [**Installation Guide**](/guide/install) — Install across Word, Excel, and PowerPoint with one PowerShell command.
