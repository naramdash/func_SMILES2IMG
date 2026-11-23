# Missing Ribbon Tab

Steps to restore the **`[SMILES]`** ribbon tab when it does not appear in Microsoft Excel, Word, or PowerPoint after installation.

## 1. Completely Restart Office

Office loads add-ins only during application startup.

1. Save your work and close **every open window** of Excel, Word, or PowerPoint.
2. Press <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>Esc</kbd> to open **Task Manager** and confirm no background `EXCEL.EXE`, `WINWORD.EXE`, or `POWERPNT.EXE` process is lingering.
3. Reopen Office and check the top ribbon for **`[SMILES]`**.

## 2. Re-Enable Disabled Add-Ins

If Office crashed previously, it may place an add-in into its **Disabled Items** list automatically.

1. Open **File** ➔ **Options** ➔ **Add-ins**.
2. In the bottom **Manage** dropdown, select **Disabled Items** ➔ click **Go...**.
3. If `func_SMILES2IMG` or `Smiles2Img` is listed, select it, click **Enable**, and restart Office.

## 3. Verify the Active Add-In Checkbox

* **Microsoft Excel:** Open **File** ➔ **Options** ➔ **Add-ins** ➔ **Manage: Excel Add-ins** ➔ **Go...**, and make sure **`func_SMILES2IMG`** (`Smiles2Img-AddIn64-packed.xll` or `Smiles2Img-AddIn-packed.xll`) is checked.
* **Microsoft Word / PowerPoint:** Open **File** ➔ **Options** ➔ **Add-ins** ➔ **Manage: COM Add-ins** ➔ **Go...**, and make sure **`func_SMILES2IMG`** is checked.

## 4. Repair Registry & Files in 10 Seconds

Re-running the [10-Second PowerShell Install](/guide/install) command safely repairs per-user `HKCU` registry registrations and refreshes binaries in `%APPDATA%\Microsoft\AddIns\Smiles2Img` without requiring administrator rights.

## Related Troubleshooting

- [**Installation Blocks**](/guide/troubleshooting/install) — Unblock Windows 11 Smart App Control or manually downloaded .xll files.
- [**Excel Formula Errors**](/excel/troubleshooting/formula-errors) — Fix #NAME?, #VALUE!, and .FLOAT legacy mode errors.
