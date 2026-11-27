# Double-Click & Re-Editing Issues

Fixes when double-clicking a Word chemical structure opens Word's **Picture Format** tab instead of **Ketcher 2D** or **Mol\* 3D**, plus best practices for keeping chemical metadata intact across co-author reviews.

## 1. Double-Click Opens Picture Format Instead of Ketcher or Mol\* {#double-click}

Every `func_SMILES2IMG` picture embeds a chemical tag inside **Alt Text** (`SMILES:...`, `KET:...`, or `MOLSTAR:...`), along with `CustomXMLParts` for Mol\* 3D sessions. Double-click falls back to Word's built-in Picture Format pane when:

* **Image came from a screenshot or PDF paste:** Plain bitmaps carry no chemical metadata. Re-insert the molecule using **`[SMILES]` ➔ `[Insert / Edit Structure]`** or **`[Paste Structure]`**.
* **Alt Text was manually cleared or stripped:** Right-click the image ➔ **View Alt Text** to confirm the `SMILES:`, `KET:`, or `MOLSTAR:` prefix is still present.
* **Word COM add-in is disabled:** If the **`[SMILES]`** tab is missing from your ribbon as well, inspect **File ➔ Options ➔ Add-ins ➔ COM Add-ins / Disabled Items** ([Missing Ribbon Tab](/guide/troubleshooting/ribbon-tab)).
* **Floating picture event was temporarily bypassed:** Switch once between text wrapping modes (*Inline* vs. *Square*) or trigger the editor directly from the ribbon. *(Note: double-click interception is scoped to the Word document canvas `_WwG` client area so ribbon or task-pane clicks are never hijacked.)*

## 2. Sharing With Mac Co-Authors & Preserving Metadata {#sharing}

* **Co-authors without the add-in:** Embedded 2D PNGs and 3D `.glb` models display and print identically on **macOS**, Windows, and Word Online—no `func_SMILES2IMG` installation needed on the reader's machine.
* **Round-trip editing survives reviews:** Because chemical tags live inside standard `.docx` **Alt Text** and `CustomXMLParts`, your colleagues can edit prose or track changes on macOS, and double-click editing still works when you reopen the file on Windows ([Document Sharing](/guide/troubleshooting/document-sharing)).
* **Vector Illustrator / TOC graphics:** Select any scheme and run **`[SMILES]` ➔ `[Copy SVG]`** to paste editable vector paths.

## 3. Alternative Ways to Re-Open the Editor

Even when double-click is overridden by another add-in or Picture Format habit, you can reopen the molecule in place at any time:

1. Select the image and click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) or **`[3D Structure (Mol*)]`** (3D) on the ribbon.
2. Or right-click the image ➔ **`Chemical Structure` ➔ `Insert / Edit Structure...`** (2D) or **`3D Structure (Mol*)...`** (3D).

![Word Chemical Structure right-click context menu](/word-context-menu-guide.png)

## Related Guides

- [**Table Clipping & Batch Style**](/word/troubleshooting/table-layout) — Fix top-clipped inline structures, shifting table widths, and `Batch Style` selection scope
- [**Ribbon & Context Menus**](/word/reference) — Complete Word `[SMILES]` ribbon map and 5 right-click context menus
- [**Missing Ribbon Tab**](/guide/troubleshooting/ribbon-tab) · [**Document Sharing**](/guide/troubleshooting/document-sharing) — Re-enable disabled Word COM add-ins and verify cross-platform files
