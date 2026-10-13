# Frequently Asked Questions (FAQ) & Troubleshooting

Find quick answers to common questions and troubleshooting tips for **SMILES for Microsoft Office Suite**.

---

## ❓ General & Excel Questions

### Q. Which formula should I use for my Excel version?
* **Microsoft 365 or Excel 2024+:** Use `=SMILES2IMG(...)` for real in-cell pictures that sit directly inside cell contents and move/scale automatically.
* **Excel 2016, 2019, 2021:** Use `=SMILES2IMG.FLOAT(...)` because older Excel builds lack native in-cell picture support. Floating images automatically move and size with cells (`xlMoveAndSize`).

### Q. The formula returns `#VALUE!`. How do I find why it failed?
1. Verify that the referenced cell contains a valid SMILES string.
2. If referencing merged cells, make sure you point to the top-left cell of the merge.
3. **Inspect Exact Chemical Errors:** From the Excel ribbon, click **[SMILES] ➔ [Show diagnostics]**. The diagnostic log displays the exact chemical syntax error (such as unclosed rings `cycle not closed`, invalid valences, or unrecognized atoms) and character position.

### Q. The molecular image looks too small.
Images scale automatically to preserve aspect ratios without clipping. Increase the cell's row height (e.g., 80–120 pt) and column width (e.g., 25–40) to allow a larger, clear rendering.

### Q. How do I match monochrome journals or dark mode themes?
* **Black & White Line Art:** Pass `"bw"` as the 3rd argument for 100% monochrome patent and ACS print style (`=SMILES2IMG(A2, "", "bw")`).
* **Dark Mode:** Pass a dark background (e.g. `"#1a1a1a"`) as the 2nd argument and white ink (`"white"`) as the 3rd argument:
  ```text
  =SMILES2IMG(A2, "#1a1a1a", "white")
  ```

---

## 🛡️ Security & Loading Issues

### Q. [Smart App Control] Excel displays "The file format and extension do not match" or formulas evaluate to `#NAME?`.
* **Cause:** Windows 11 **Smart App Control (SAC)** blocks unsigned open-source XLL add-in binaries when enforcement mode is active.
* **Important:** Disabling Smart App Control does **NOT** disable your antivirus; **real-time Microsoft Defender protection remains 100% active**.
* **Fix (10 seconds):**
  1. Open Windows **Settings** (`Win + I`) ➔ **Privacy & security** ➔ **Windows Security**.
  2. Click **App & browser control** ➔ **Smart App Control settings**.
  3. Set to **Off** and restart Excel.

### Q. Excel blocks or refuses to load the XLL add-in.
Web downloads are marked by Windows with a `Zone.Identifier` stream. Right-click the `.xll` file, select **Properties**, check **Unblock** at the bottom, and click **OK**.
Alternatively, run the [One-Line Quick Install](../guide/getting-started) command, which removes security blocks automatically.
