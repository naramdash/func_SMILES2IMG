using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Smiles2Img.COM;
using Smiles2Img.Rendering;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Smiles2Img.PowerPoint;

[ComVisible(true)]
[Guid("A755973C-21FB-494A-A3C5-086AFFB7C2AC")]
[ProgId("Smiles2Img.PowerPoint.AddIn")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class PowerPointAddIn : IDTExtensibility2, IRibbonExtensibility
{
    private MSPowerPoint.Application? pptApp;
    private MSPowerPoint.EApplication_Event? pptEvents;
    private IRibbonUI? ribbon;

    private const int WH_MOUSE = 7;
    private const int WM_LBUTTONDBLCLK = 0x0203;

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    private HookProc? mouseHookProc;
    private IntPtr hMouseHook = IntPtr.Zero;
    private bool isEditing = false;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            pptApp = (MSPowerPoint.Application)Application;
            var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            NativeIndigo.Initialize(asmDir);

            pptEvents = (MSPowerPoint.EApplication_Event)pptApp;
            pptEvents.WindowBeforeDoubleClick += PowerPointStructureManager.OnWindowBeforeDoubleClick;

            // PowerPoint's native WindowBeforeDoubleClick event does not fire for shapes/pictures in Normal view.
            // Install an in-process thread-level mouse hook to intercept double-clicks on molecular structure shapes.
            try
            {
                mouseHookProc = MouseHookCallback;
                hMouseHook = SetWindowsHookEx(WH_MOUSE, mouseHookProc, IntPtr.Zero, GetCurrentThreadId());
            }
            catch { }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to initialize SMILES for PowerPoint:\n" + ex.Message, "SMILES2IMG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        try
        {
            if (hMouseHook != IntPtr.Zero)
            {
                try { UnhookWindowsHookEx(hMouseHook); } catch { }
                hMouseHook = IntPtr.Zero;
            }
            mouseHookProc = null;

            if (pptEvents != null)
            {
                pptEvents.WindowBeforeDoubleClick -= PowerPointStructureManager.OnWindowBeforeDoubleClick;
                pptEvents = null;
            }
            pptApp = null;
        }
        catch { }
    }

    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDBLCLK && !isEditing && pptApp != null)
        {
            try
            {
                var sel = pptApp.ActiveWindow?.Selection;
                if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    var targetShape = sel.ShapeRange[1];
                    var structStr = PowerPointStructureManager.ExtractStructure(targetShape);
                    if (!string.IsNullOrEmpty(structStr))
                    {
                        isEditing = true;
                        try
                        {
                            PowerPointStructureManager.EditSelectedMolecule(pptApp, targetShape);
                        }
                        finally
                        {
                            isEditing = false;
                        }
                        return (IntPtr)1; // Suppress PowerPoint's default double-click action (e.g. Picture Format pane)
                    }
                }
            }
            catch { }
        }
        return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
    }

    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public string GetCustomUI(string RibbonID)
    {
        return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabSmiles"" label=""SMILES"">
        <group id=""grpMolecules"" label=""Chemical Structures"">
          <button id=""btnInsertMolecule""
                  label=""Insert Structure""
                  screentip=""EPAM Ketcher Molecule Editor""
                  supertip=""Open Ketcher web editor to draw chemical structures and insert into the current slide.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnQuickInsert""
                  label=""Quick Insert""
                  screentip=""Quick Insert from Clipboard or Selection""
                  supertip=""Convert SMILES or KET directly into a molecular structure image without opening the editor.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
          <button id=""btnEditMolecule""
                  label=""Edit Structure""
                  screentip=""Edit Selected Structure""
                  supertip=""Load the selected molecule image back into the Ketcher editor for editing.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
          <button id=""btnCopySmiles""
                  label=""Copy SMILES""
                  screentip=""Copy SMILES of Selected Molecule""
                  supertip=""Extract SMILES string from the selected molecule and copy it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""Copy Molecule Info""
                  screentip=""Copy Formula and Molecular Weight""
                  supertip=""Calculate and copy chemical formula, molecular weight (MW), and exact mass to clipboard.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPptPictureEdit""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopyInfo""
              label=""Copy Molecule Info (Formula/MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPptPicture"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextPptShapeEdit""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopyInfo""
              label=""Copy Molecule Info (Formula/MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPptShape"" insertBeforeMso=""Cut"" />
    </contextMenu>
  </contextMenus>
</customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        ribbon = ribbonUI;
    }

    public void OnInsertMolecule(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.InsertMolecule(pptApp);
    }

    public void OnQuickInsert(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.QuickInsertFromClipboard(pptApp);
    }

    public void OnEditMolecule(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.EditSelectedMolecule(pptApp);
    }

    public void OnCopySmiles(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeSmiles(pptApp);
    }

    public void OnCopyKet(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeKet(pptApp);
    }

    public void OnCopyMoleculeInfo(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeInfo(pptApp);
    }
}
