# 3. 3D Protein Figures (`Mol*`)

Embed 3D protein, nucleic acid, and ligand-pocket figures directly inside your Microsoft Word manuscript—fetching public entries by 4-character PDB ID or loading local `.pdb` / `.cif` files offline, with saved camera orientation for re-editing.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## 1. Insert a Protein or Binding-Pocket Figure

1. Place your cursor where the figure belongs in your Word manuscript.
   ::: tip Power Tip: Preload a PDB Code Straight From Your Draft
   If your manuscript text already mentions a PDB ID like `1CRN` or `6LU7` (or a local `.pdb`/`.cif` file path), highlight that text **before** clicking **`[3D Structure (Mol*)]`**—Mol\* automatically fetches and loads that exact entry on open.
   :::
2. Click **`[SMILES]` ➔ `[3D Structure (Mol*)]`** on the Word ribbon (or right-click ➔ **`Chemical Structure` ➔ `3D Structure (Mol*)...`**).
3. Load a structure in either of two ways:
   * **Public PDB Fetch:** Type a 4-character PDB ID (e.g., `1CRN`, `4HHB`, `6LU7`, `1BNA`) and click **`Fetch`**.
   * **Offline Local File:** Click **`Open File...`** to load `.pdb`, `.cif`, `.bcif`, `.ent`, `.mol`, `.sdf`, `.molj`, or `.molx` files without sending data over the network.
4. Adjust your representation (**Cartoon**, **Surface**, **Ball & Stick** pocket residues), coloring, and background (**`White`** or **`Transparent`**), select your output mode on the top bar, and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**).

## 2. Choose Between `2D Picture` and `3D Model (GLB)`

| Word Output Mode | How It Behaves in `.docx` | Best Manuscript Use Case |
| :--- | :--- | :--- |
| **`2D Picture` (Recommended for Papers)** | Embeds a high-DPI `.png` snapshot (`BG: White` or `Transparent`) | **ACS / Nature / Elsevier submissions, PDF conversion, and print manuscripts** across all Word versions |
| **`3D Model (GLB)`** | Embeds a native interactive 3D object users can rotate 360° inside the Word page | **Internal lab reports, grant review drafts, and interactive `.docx` reviews** (requires Word 2019 / 2021 / 2024 / Microsoft 365) |

::: info Choosing Background & Contrast for Word Figures
For standard white journal pages and PDF figure plates, **`BG: White`** or **`BG: Transparent`** produces clean edges without gray halos. Use **`BG: Transparent`** whenever placing a protein inset inside a shaded callout box or multi-panel figure layout.
:::

## 3. Re-Editing Camera Angles & Binding-Pocket Views

* **Exact View Preservation:** `func_SMILES2IMG` compresses and embeds the complete Mol\* viewer state (PDB target, camera coordinates, zoom, representation, and colors) inside Word's native **`CustomXMLParts`**, linked via the shape's `MOLSTAR:` Alt Text tag.
* **Refinement in Place:** **Double-click** any inserted Mol\* 2D figure or 3D model in Word to reopen Mol\* at the exact camera orientation you left off—tweak rotation slightly or toggle ligand surfaces and apply in place without rebuilding the view from scratch.
* **Protected From `Batch Style`:** Document-wide **`[Batch Style]`** (such as switching all 2D schemes to `Publication B&W`) automatically skips Mol\* 3D figures so your colored protein ribbons and surfaces stay intact.

## Related Guides

- [**Mol\* 3D Viewer Manual**](/editor/molstar) — Sample PDB IDs (`1CRN`, `4HHB`, `6LU7`, `1BNA`), representation controls, and offline PDB loading
- [**4. Journal Finish, SVG, Sharing**](/word/publishing) — Finalize 2D schemes with ACS B&W while keeping Mol\* figures untouched
- [**Offline PDB Files Troubleshooting**](/editor/troubleshooting/offline-pdb) — Open local `.pdb` / `.cif` files behind corporate firewalls
