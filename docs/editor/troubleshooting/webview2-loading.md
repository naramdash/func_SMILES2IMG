# WebView2 Loading

Fixes when the embedded **Ketcher 2D** sketcher or **Mol\* 3D** viewer window opens blank or fails to initialize.

## 1. Verify Microsoft Edge WebView2 Runtime

Both Ketcher 2D and Mol\* 3D run locally inside **Microsoft Edge WebView2**.

* Windows 10 and Windows 11 include Edge WebView2 by default, but stripped enterprise OS images occasionally omit it.
* If the editor dialog cannot initialize WebView2, install the official **Microsoft Edge WebView2 Evergreen Runtime** and restart Office.

## 2. Reset the Local Editor Bundle Cache

`func_SMILES2IMG` extracts its offline Ketcher and Mol\* assets into `%LOCALAPPDATA%\Smiles2Img` and validates them using a `SHA-256` fingerprint. If an antivirus lock or interrupted update left partial files:

1. Close Excel, Word, and PowerPoint completely.
2. Press <kbd>Win</kbd> + <kbd>R</kbd>, type `%LOCALAPPDATA%\Smiles2Img`, and press <kbd>Enter</kbd>.
3. Delete the extracted cache subfolders inside `%LOCALAPPDATA%\Smiles2Img`.
4. Reopen Office and launch Ketcher or Mol\*—fresh verified bundles are extracted automatically.

## 3. High-DPI Scaling or Off-Screen Window

Editor dialogs automatically scale to your monitor's DPI (`100%`–`200%`) and clamp to your active screen working area. If you unplugged a secondary monitor while a dialog was open, press <kbd>Alt</kbd> + <kbd>Space</kbd> ➔ **M** and use arrow keys, or close and reopen the editor from the `[SMILES]` ribbon.

## Related Guides

- [**Offline PDB Files**](/editor/troubleshooting/offline-pdb) — Open local .pdb, .cif, and .sdf files in Mol* behind strict firewalls.
- [**Ketcher 2D Sketcher Guide**](/editor/) — Shortcuts, reaction schemes, and background modes across Office.
