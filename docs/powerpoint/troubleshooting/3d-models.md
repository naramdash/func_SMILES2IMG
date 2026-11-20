# 3D Models & Editing

Fixes for rotatable 3D `.glb` protein models, PowerPoint double-click editing, and slide deck sharing.

## 1. 3D `.glb` Model Does Not Insert or Rotate

* **Office Version Requirement:** Interactive 3D `.glb` models require **PowerPoint 2019, 2021, 2024, or Microsoft 365**.
* **Older PowerPoint (2016) or Static PDF Handouts:** Switch the Mol\* top toolbar from **`3D Model (GLB)`** to **`2D Picture`** to insert a high-resolution PNG figure (transparent or solid background) that works on every PowerPoint version.

## 2. Double-Click Does Not Reopen Ketcher or Mol\* on a Slide

`func_SMILES2IMG` checks each slide shape's **Alt Text** (`SMILES:...`, `KET:...`, or `MOLSTAR:...`) and deck `CustomXMLParts` to reopen the exact 2D canvas or 3D camera angle.

* **Plain Screenshots:** Pasted bitmaps without chemical metadata open PowerPoint's standard Format Picture pane instead.
* **Verify Alt Text:** Right-click the shape ➔ **View Alt Text** to confirm the metadata tag is still present.
* **Shortcut Alternative:** Select the shape and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (2D) or click **`[SMILES]` ➔ `[3D Structure (Mol*)]`** (3D).
* **COM Add-In Check:** If the `[SMILES]` tab itself is missing, see [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab).

## 3. Presenting on Another PC or Mac Without the Add-In

Inserted 2D PNGs and 3D `.glb` models are standard native PowerPoint objects—your slides present and rotate 3D models cleanly on conference PCs or **macOS** even when `func_SMILES2IMG` is not installed. See [Document Sharing](/guide/troubleshooting/document-sharing).

## Related Guides

- [**Dark Slide Backgrounds**](/powerpoint/troubleshooting/dark-slides) — Remove white boxes and switch decks to transparent white-ink Dark Mode.
- [**Offline PDB Files (Mol\*)**](/editor/troubleshooting/offline-pdb) — Load local .pdb / .cif structures when firewalls block RCSB PDB fetches.
