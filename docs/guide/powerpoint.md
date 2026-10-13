# SMILES for Microsoft PowerPoint

**SMILES for Microsoft PowerPoint** brings effortless chemical presentations to life with one-click slide insertion, dark theme transparent backgrounds, and seamless in-place editing.

---

## 🎥 3-Minute Video Walkthrough

> 💡 **Master chemical slide design in under 2 minutes (1080p high definition).**  
> Covers slide center auto-placement, transparent background toggles for dark templates, 0.1s text box substitution, aspect-ratio preserving double-click edits, and multi-step reaction equations.

<p align="center">
  <video controls="controls" width="100%" style="max-width: 800px; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.15);" poster="/powerpoint-showcase.png">
    <source src="/videos/powerpoint-guide-video.mp4" type="video/mp4" />
    <source src="https://github.com/user-attachments/assets/b7ac8c3f-cc06-4820-8df6-1593d07ccb97" type="video/mp4" />
    <a href="/videos/powerpoint-guide-video.mp4">🎬 Watch PowerPoint 3-Minute Video Walkthrough</a>
  </video>
</p>

<details>
<summary><b>⏱️ Video Chapters & Timeline</b></summary>

| Time | Chapter | Highlights |
| :---: | :--- | :--- |
| **00:00** | **Overview & Features** | One-click slide insertion, transparent/white backgrounds, double-click edits, and dark slide support |
| **00:13** | **Step 1: Insert on Slide** | Click `[Insert Structure]` on the ribbon to auto-place high-res vector-sharp PNGs directly onto the active slide |
| **00:28** | **Step 2: Transparent Background** | Toggle `[Background: Transparent]` in Ketcher for flawless presentation against dark themes |
| **00:43** | **Step 3: Quick Replace** | Select any slide text box or bullet containing SMILES and click `[Quick Insert]` for instant shape conversion |
| **00:58** | **Step 4: Double-Click In-Place** | Double-click the molecule shape on the slide to edit without shifting its position or altering aspect ratios |
| **01:14** | **Step 5: Reactions & Properties** | Multi-step organic reaction pathways stored losslessly in `Shape.Tags['KET']`; instant Q&A molecular weight calculation |
| **01:29** | **Summary & Checklist** | Conference presentation, seminar, and lecture slide best practices |

</details>

---

## 🌟 Core Highlights

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" />
</p>

::: tip Key Advantages in PowerPoint
* **Dark Theme Optimization:** Switch between crisp white backgrounds and alpha-transparent PNGs at any time so chemical structures blend seamlessly into modern dark slide decks.
* **Exact Geometric Preservation:** When double-clicking to edit, the new structure retains the exact slide coordinates (`Left`, `Top`), width, and center anchor point.
* **Lossless Slide Portability:** Stored inside PowerPoint shapes using `Shape.Tags['KET']` and `Shape.Tags['SMILES']`.
:::

---

## 📑 Ribbon Menu: `[SMILES]` Tab

1. **Insert Structure:** Launches Ketcher. Draws molecule or reaction, then clicks **Apply** to place at the center of the active slide.
2. **Quick Insert:**
   - If a slide shape or text frame is selected, converts the text into a molecule shape in 0.1s.
   - If no shape is selected, pastes from clipboard directly onto the slide.
3. **Insert from Clipboard:** Reads clipboard and inserts onto the current slide.
4. **Settings & Diagnostics:** Engine health and log viewer.

---

## 🌓 Dark Theme & Transparent Backgrounds

1. When editing or inserting via Ketcher, look at the top status bar.
2. Toggle between **Background: White** and **Background: Transparent**.
3. Transparent mode exports clean alpha-channel PNGs, ensuring no awkward white rectangular borders appear around your chemical bonds on dark backgrounds.

---

## 🖱️ Double-Click In-Place Editing

1. Double-click any chemical shape on the presentation slide.
2. The embedded Ketcher editor immediately opens with all atoms, bonds, and reaction arrows intact.
3. Make changes and press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
4. The shape updates on the slide instantly, retaining its exact slide position and proportional scale.

---

## 📋 Context Menu & Presentation Q&A

During slide preparation or seminar Q&A, right-click the chemical structure:

* **Chemical Structure ➔ Copy Molecular Weight (MW):** Instant g/mol value for audience questions.
* **Chemical Structure ➔ Copy Molecular Formula:** Formula (e.g., `C8H10N4O2`).
* **Chemical Structure ➔ Copy SMILES / Copy KET:** Full text representation.
