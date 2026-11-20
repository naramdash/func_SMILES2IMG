# Offline PDB Files

How to load 3D macromolecular structures in **Mol\*** when a corporate firewall blocks direct RCSB PDB / PDBe downloads.

## 1. Why 4-Letter PDB Code Fetches Can Fail

Entering a 4-character PDB code (such as `1CRN`, `4HHB`, or `6LU7`) and clicking **`Fetch`** downloads the public crystal structure over HTTPS from **RCSB PDB / PDBe**.

If your lab or enterprise network restricts external HTTPS endpoints, the remote fetch will time out or fail.

## 2. Load Local 3D Structure Files 100% Offline

The Mol\* 3D viewer runtime itself is bundled locally inside `func_SMILES2IMG`, so local structure files require **zero internet access**:

1. Obtain the structure file locally (`.pdb`, `.cif`, `.bcif`, `.sdf`, `.mol`, or `.mol2`).
2. Click **`[SMILES]` ➔ `[3D Structure (Mol*)]`** on the ribbon.
3. Click **`Open File...`** in the top Mol\* toolbar and select your local file.
4. Style the cartoon/surface view as usual and click **`Apply to Slide`** / **`Apply to Document`** (or press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>).

## 3. Re-Editing Offline Models Later

Your camera angle, style recipe (`.molj`), and model reference persist inside the Office document metadata (`CustomXMLParts`), so double-clicking reopens your exact view later.

## Related Guides

- [**WebView2 Loading**](/editor/troubleshooting/webview2-loading) — Fix blank Ketcher/Mol* windows and reset %LOCALAPPDATA%\Smiles2Img cache.
- [**Mol\* 3D Protein Viewer Guide**](/editor/molstar) — 3-step walkthrough for 3D GLB models and publication 2D protein renders.
