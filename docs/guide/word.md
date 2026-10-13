# SMILES for Microsoft Word

**SMILES for Microsoft Word** brings seamless chemical structure editing and inline high-resolution molecule rendering directly into Word documents, research papers, and thesis templates.

---

## 🎥 3-Minute Video Walkthrough

> 💡 **Master chemical structure editing in Microsoft Word in under 2 minutes (1080p high definition).**  
> Covers ribbon controls, 0.1s Quick Insert text substitution, table cell layout preservation during double-click editing, lossless reaction equation storage, and molecular weight calculation.

<p align="center">
  <video src="https://github.com/user-attachments/assets/e34fc58a-b320-480e-8e83-876356796b6e" controls="controls" width="100%" style="max-width: 800px; border-radius: 8px;">
    <a href="https://github.com/user-attachments/assets/e34fc58a-b320-480e-8e83-876356796b6e">🎬 Watch Word 3-Minute Video Walkthrough</a>
  </video>
</p>

<details>
<summary><b>⏱️ Video Chapters & Timeline</b></summary>

| Time | Chapter | Highlights |
| :---: | :--- | :--- |
| **00:00** | **Overview & Features** | Inline high-res shapes, double-click in-place editing, 0.1s Quick Insert, and lossless KET reactions |
| **00:13** | **Step 1: Insert Structure** | Click `[Insert Structure]` on the `[SMILES]` ribbon tab to draw in Ketcher and insert into the document |
| **00:28** | **Step 2: Quick Insert** | Highlight any SMILES string or copy to clipboard, then click `[Quick Insert]` for instant conversion |
| **00:43** | **Step 3: Double-Click Editing** | Double-click the inline structure inside paragraphs or table cells to re-edit in Ketcher without disrupting layout |
| **00:59** | **Step 4: Lossless Reactions** | Reaction arrows, plus signs, and text annotations are permanently preserved in document metadata (`AlternativeText`) |
| **01:14** | **Step 5: Molecular Properties** | Right-click the image to calculate chemical formula and molecular weight (g/mol) |
| **01:28** | **Summary & Checklist** | Essential tips for research manuscripts, journal submissions, and theses |

</details>

---

## 🌟 Core Highlights

<p align="center">
  <img src="/word-showcase.png" alt="SMILES for Word Showcase" width="100%" />
</p>

::: tip Key Advantages in Word
* **Inline Picture Flow:** Structures are inserted as native `InlineShapes` that behave like characters. They wrap naturally with sentences, align within table cells, and move with paragraphs.
* **Layout Preservation:** When double-clicking to edit a structure, the newly generated image replaces the old shape in-place without distorting table column widths or paragraph margins.
* **Lossless File Portability:** All structure metadata (`KET` or `SMILES`) is stored inside the `.docx` XML structure (`AlternativeText` and `Title`). Sharing the file with colleagues keeps all chemical metadata intact.
:::

---

## 📑 Ribbon Menu: `[SMILES]` Tab

<p align="center">
  <img src="/ribbon-guide.png" alt="Word Ribbon Menu" width="100%" />
</p>

1. **Insert Structure:** Launches the offline Ketcher 2D editor. Draw your molecule or reaction, then press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (or click **Apply**) to insert at the current cursor position.
2. **Quick Insert:**
   - If text is selected in your document, converts that SMILES text directly into an inline chemical structure picture in 0.1 seconds.
   - If no text is selected, parses the clipboard and inserts the structure at the cursor.
3. **Insert from Clipboard:** Reads clipboard content as SMILES or KET and embeds it as an inline shape.
4. **Settings & Diagnostics:** Displays technical diagnostics, version details, and troubleshooting logs.

---

## 🖱️ Double-Click In-Place Editing

1. Double-click any chemical structure picture inside your Word document.
2. The embedded Ketcher editor instantly opens with the exact molecule or reaction graph loaded.
3. Modify bonds, functional groups, or annotations.
4. Click **[Apply (Ctrl+Enter)]**. The image updates in-place immediately, retaining its original alignment and scaling.

---

## 📋 Context Menu Options

Right-click any chemical structure image or text inside Word:

* **Chemical Structure ➔ Edit Structure in Ketcher:** Opens Ketcher editor.
* **Chemical Structure ➔ Copy SMILES / Copy KET:** Copies clean representation to clipboard.
* **Chemical Structure ➔ Copy Molecular Formula:** Copies Hill-notation formula (e.g. `C8H10N4O2`).
* **Chemical Structure ➔ Copy Molecular Weight (MW):** Copies accurate molecular weight in g/mol.
