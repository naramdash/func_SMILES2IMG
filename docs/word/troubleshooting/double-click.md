# Double-Click Editing

Why double-clicking a chemical structure in Microsoft Word might not reopen **Ketcher 2D** or **Mol\* 3D**—and how to restore it.

## 1. How Word Recognizes Editable Structures

Every structure inserted by `func_SMILES2IMG` saves chemical metadata inside the picture's **Alt Text** (`SMILES:...`, `KET:...`, or `MOLSTAR:...`) and lossless coordinates/recipes inside Word `CustomXMLParts`.

Both **inline pictures** (`InlineShape`) and **floating pictures / 3D models** (`Shape`) support double-click re-editing in place.

## 2. Common Reasons Double-Click Does Not Open the Editor

| Cause | Why It Happens | Fix |
| :--- | :--- | :--- |
| **Plain Screenshot or External Paste** | An image pasted from Snipping Tool, a PDF, or another app carries no chemical metadata. | Insert or paste via **`[SMILES]` ➔ `[Insert / Edit Structure]`** or **`[Paste Structure]`**. |
| **Alt Text Was Cleared** | Another macro or manual edit deleted the `SMILES:`, `KET:`, or `MOLSTAR:` prefix. | Right-click the picture ➔ **View Alt Text** to confirm the tag is intact. |
| **Word COM Add-In Disabled** | Double-click interception requires the Word COM add-in to be loaded. | Check [Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab) (`COM Add-ins` / `Disabled Items`). |

## 3. Shortcut & Ribbon Alternatives

Select any chemical picture in Word and use:

* <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> to open Ketcher 2D immediately
* **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) or **`[3D Structure (Mol*)]`** (3D)
* **Right-click** the structure ➔ **`Chemical Structure` ➔ `Edit Structure...`**

## Related Guides

- [**Table Layout & Styling**](/word/troubleshooting/table-layout) — Preserve table column widths on re-edit and apply ACS B&W Batch Style cleanly.
- [**Document Sharing**](/guide/troubleshooting/document-sharing) — How Word .docx structures display on Mac and PCs without the add-in.
