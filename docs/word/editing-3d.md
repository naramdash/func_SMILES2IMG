# In-Place Editing & 3D Proteins

Re-edit any inserted 2D structure in place without disturbing Word table widths, or embed interactive 3D protein models and high-DPI 2D renders via **Mol\***.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## 1. Double-Click In-Place 2D Editing

* **Double-click** any chemical picture in prose or inside a table cell to reopen it in Ketcher 2D in `<0.05s`.
* When you apply changes, `func_SMILES2IMG` replaces the graphic in place with aspect ratio locked (`LockAspectRatio = msoTrue`) so table columns and paragraph spacing stay intact.

## 2. Insert 3D Protein Figures (`Mol*`)

1. Click **`[SMILES]` ➔ `[3D Structure (Mol*)]`** on the ribbon.
2. Enter a 4-letter PDB code (such as `1CRN` or `6LU7`) and click **`Fetch`**, or open a local `.pdb` / `.cif` file.
3. Style the model (Cartoon, Surface, Ball & Stick) and choose:
   - **`3D Model (GLB)`** — interactive 360°-rotatable model inside Word
   - **`2D Picture`** — high-DPI publication `.png` figure (White, Transparent, or Dark background)
4. Double-click any inserted 3D model or 2D render later to reopen Mol\* at the exact same camera angle.
