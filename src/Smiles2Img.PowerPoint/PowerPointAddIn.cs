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
            MessageBox.Show("SMILES for PowerPoint 초기화 실패:\n" + ex.Message, "SMILES2IMG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    var smiles = PowerPointStructureManager.ExtractSmiles(targetShape);
                    if (!string.IsNullOrEmpty(smiles))
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
      <tab id=""tabSmiles"" label=""SMILES"" insertAfterMso=""TabInsert"">
        <group id=""grpMolecules"" label=""분자 구조식 (Structure)"">
          <button id=""btnInsertMolecule""
                  label=""분자 구조식 삽입""
                  screentip=""EPAM Ketcher 분자 에디터""
                  supertip=""Ketcher 웹 에디터를 열어 화학 구조식을 그리고 현재 슬라이드에 삽입합니다.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnEditMolecule""
                  label=""구조식 편집""
                  screentip=""선택된 분자 구조식 편집""
                  supertip=""선택된 분자 이미지를 Ketcher 에디터로 다시 불러와 수정합니다.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
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

    public void OnEditMolecule(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.EditSelectedMolecule(pptApp);
    }
}
