# Ketcher 2D Molecular & Reaction Editor

**SMILES for Microsoft Office Suite** includes an embedded, fully offline **EPAM Ketcher v3.18.0** editor powered by Microsoft Edge WebView2.

---

## 🌟 Overview & Capabilities

<p align="center">
  <img src="/ketcher-guide.png" alt="Ketcher Editor Interface" width="100%" />
</p>

* **100% Offline & Private:** Operates entirely locally within your machine. No web connections, external telemetry, or cloud servers are accessed.
* **Full Chemical Drawing Arsenal:** Single, double, triple, and aromatic bonds; stereochemical wedges/hashes; heteroatoms; charges; valence tools; and functional group templates.
* **Organic Reaction Mapping:** Multi-step reaction schemes with curved/straight reaction arrows, plus signs, and reaction condition text.
* **Real-Time Property Calculation:** Automatically computes the molecular formula (Hill notation), molecular weight (MW in g/mol), and exact monoisotopic mass as you draw.

---

## ⌨️ Essential Keyboard Shortcuts

| Shortcut | Action | Description |
| :--- | :--- | :--- |
| <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> | **Open Editor (Excel)** | Opens Ketcher for the currently selected cell |
| <kbd>Ctrl</kbd> + <kbd>Enter</kbd> | **Apply to Document** | Renders drawing into active cell/document and closes editor |
| <kbd>Esc</kbd> | **Cancel** | Closes Ketcher without modifying the document |
| **Double-Click** | **Edit In-Place** | Double-click any cell or chemical image to edit in Ketcher |

---

## 🎨 Workflow & Background Transparency

<p align="center">
  <img src="/workflow-ketcher.png" alt="Ketcher Workflow" width="100%" />
</p>

### 1. Launching Ketcher
* **Excel:** Select a cell and press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>, or right-click ➔ **Chemical Structure ➔ Edit Structure in Ketcher**.
* **Word / PowerPoint:** Click **[Insert Structure]** on the `[SMILES]` ribbon tab, or double-click an existing structure.

### 2. Drawing Molecules & Reactions
* Use the left toolbar to select bond types and templates (benzene, cyclopentane, rings, amino acids).
* Use the right toolbar to choose periodic table elements (C, H, N, O, P, S, halogens).
* Add reaction arrows and plus signs using the bottom-left reaction tool.

### 3. Background Transparency Toggle
In the Ketcher top toolbar:
* **Background: White:** Best for standard white documents, printed papers, and classic Excel spreadsheets.
* **Background: Transparent:** Exports pure alpha-channel PNGs with transparent backgrounds, ideal for dark PowerPoint slides or colored table cells.

### 4. Applying to Document
Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**). The molecule or reaction canvas is immediately rasterized using EPAM Indigo and placed precisely into your document.
