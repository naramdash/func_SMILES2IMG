# SMILES for Microsoft PowerPoint

**SMILES for Microsoft PowerPoint** brings effortless chemical presentations to life with one-click slide insertion, dark theme transparent backgrounds, and seamless in-place editing.

---

## 🌟 Visual Feature Walkthrough

Crafting scientific slides, seminar presentations, and conference decks with chemical accuracy has never been easier:

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip Key Advantages in PowerPoint
* **Dark Theme Optimization:** Switch between crisp white backgrounds and alpha-transparent PNGs at any time so chemical structures blend seamlessly into modern dark slide decks without white rectangular halos.
* **Exact Geometric Preservation:** When double-clicking to edit, the new structure retains the exact slide coordinates (`Left`, `Top`), dimensions, and center anchor point.
* **Lossless Slide Portability:** Chemical metadata is stored directly in PowerPoint shape tags (`Shape.Tags['KET']` and `Shape.Tags['SMILES']`), allowing seamless cross-machine editing.
:::

---

## 📑 Dedicated Office Ribbon: `[SMILES]` Tab

The custom `[SMILES]` tab on PowerPoint's ribbon provides one-click slide insertion and clipboard parsing:

<p align="center">
  <img src="/ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ Insert Structure (`Ctrl+Shift+K`):** Launches Ketcher. Draw your molecule or multi-step reaction, then click **Apply** (<kbd>Ctrl</kbd> + <kbd>Enter</kbd>) to place it directly at the center of the active slide.
2. **⚡ Quick Insert:**
   - If a text box or shape containing SMILES text is selected on the slide, converts the text frame into a high-resolution molecule shape in 0.1s.
   - If no shape is selected, automatically reads chemical text from the clipboard and places it onto the slide.
3. **📋 Insert from Clipboard:** Reads clipboard SMILES/KET and embeds it as a vector-crisp slide image.
4. **🔍 Settings & Diagnostics:** Opens engine diagnostic logs and troubleshooting metrics.

---

## 🖱️ Redesigned Context Menu: `Chemical Structure`

Right-click any chemical shape or blank presentation slide area:

<p align="center">
  <img src="/context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Case A: Right-Clicking an Existing Chemical Shape
* **✏️ Edit Structure in Ketcher:** Opens Ketcher with the current molecule or reaction loaded.
* **📋 Copy SMILES / Copy Molfile (KET):** Copies clean structure notation to the clipboard.
* **🔬 Copy Molecular Formula:** Copies standard Hill-notation formula (e.g. `C8H10N4O2`).
* **⚖️ Copy Molecular Weight (MW):** Instant g/mol value for conference audience Q&A.
* **🎯 Copy Exact Mass:** Copies high-resolution monoisotopic mass (e.g. `194.0804`).

### Case B: Right-Clicking Blank Slide Canvas
* **✏️ Insert Structure in Ketcher:** Launches Ketcher to draw and place a new molecule onto the slide.
* **📋 Paste from Clipboard:** Reads chemical notation from clipboard and creates a slide shape.

---

## 🧪 Ketcher 2D Molecule & Reaction Editor

Double-click any molecule shape on your slide or press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> to launch Ketcher:

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **🌓 Dark Theme & Transparent Backgrounds:** Toggle between **Background: White** and **Background: Transparent** in the top status bar. Transparent mode exports clean alpha-channel PNGs, ensuring no awkward white rectangular borders appear around your chemical bonds on dark backgrounds.
* **Live Molecular Property Banner:** Live updates for Formula, MW, and Exact Mass.
* **Hotkeys:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to update the slide instantly, <kbd>Esc</kbd> to cancel.

---

## 🔬 Multi-Step Reactions & Synthesis Schemes

Present complex chemical mechanisms, catalytic cycles, and organic synthesis pathways directly on your slides:

<p align="center">
  <img src="/reaction-scheme-guide.png" alt="Reaction Schemes on PowerPoint Slides" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **Reaction Schemes Preserved:** Full reaction schemes with arrows, catalysts, and reagents are stored losslessly inside `Shape.Tags['KET']`.
* **In-Place Re-editing:** Double-clicking the slide shape reopens the reaction canvas in Ketcher, preserving all arrows and conditions for quick layout adjustments before presentations.

---

## 🔍 Built-In Diagnostics Console

Click **[SMILES] ➔ [Settings & Diagnostics]** to view system health and engine logs:

<p align="center">
  <img src="/diagnostics-guide.png" alt="SMILES Diagnostics Console in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* Confirms PowerPoint COM Add-in registration and Indigo engine readiness.
* Verifies offline Ketcher assets and rasterization performance metrics.
