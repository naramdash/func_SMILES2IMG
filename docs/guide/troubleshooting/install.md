# Installation Blocks

Fixes for Windows 11 Smart App Control warnings, blocked `.xll` / `.zip` downloads, and restricted corporate network installations.

## 1. Windows 11 Smart App Control (SAC) Warning

If Windows 11 shows **"The file format and extension do not match"** or Excel displays `#NAME?` right after installing:

* **Cause:** Windows 11 **Smart App Control** blocks unsigned open-source `.xll` Excel add-ins when set to strict enforcement.
* **Security Note:** Turning off Smart App Control **does not disable Microsoft Defender Antivirus**—real-time virus protection stays 100% active.
* **10-Second Fix:**
  1. Open Windows **Settings** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **Privacy & security** ➔ **Windows Security**.
  2. Click **App & browser control** ➔ **Smart App Control settings**.
  3. Switch Smart App Control to **Off**, then **completely close and reopen** Office.

## 2. Manually Downloaded `.xll` or `.zip` Blocked by Windows

Files downloaded through a web browser carry a Windows Mark-of-the-Web security block flag.

* **Manual Fix:** Right-click the `.zip` archive (before extracting) or `.xll` file ➔ **Properties** ➔ check **Unblock** at the bottom ➔ click **OK**.
* **Automatic Alternative:** Run the [10-Second PowerShell Install](/guide/install) command, which calls `Unblock-File` automatically on every installed file.

## 3. Corporate Firewall or Offline Workstation

* **Installing Without Direct GitHub Access:** Follow the [Offline Manual ZIP Install](/guide/install#offline-zip) procedure.
* **What Works Offline:** Core 2D rendering (`=SMILES2IMG`, `=KET2IMG`, `=INCHI2IMG`), all physicochemical/3D/SAR calculations, built-in solvent/drug/CAS dictionary, Ketcher 2D editing, `.sdf` import, Batch Style, and SVG copy run 100% locally.
* **What Requires Internet:** Uncached PubChem name/CAS lookups (`=NAME2IMG` for rare compounds), fetching public 4-letter PDB IDs in Mol\*, and clicking **`[Check for Updates]`**.

## Related Troubleshooting

- [**Missing Ribbon Tab**](/guide/troubleshooting/ribbon-tab) — Recover a hidden or disabled [SMILES] ribbon tab in Excel, Word, or PowerPoint.
- [**Excel Formula Errors**](/excel/troubleshooting/formula-errors) — Resolve #NAME?, #VALUE!, #N/A, and .FLOAT compatibility issues.
