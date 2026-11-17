using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Smiles2Img.COM;
using Smiles2Img.Editor;
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
    private ThreadMouseDoubleClickHook? mouseHook;
    private bool isEditing = false;

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            pptApp = (MSPowerPoint.Application)Application;
            try
            {
                if (AddInInst != null)
                    ((dynamic)AddInInst).Object = this;
            }
            catch { }
            var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            NativeIndigo.Initialize(asmDir);

            pptEvents = (MSPowerPoint.EApplication_Event)pptApp;
            pptEvents.WindowBeforeDoubleClick += PowerPointStructureManager.OnWindowBeforeDoubleClick;

            // PowerPoint's native WindowBeforeDoubleClick event does not fire for shapes/pictures in Normal view.
            // Install an in-process thread-level mouse hook to intercept double-clicks on molecular structure shapes.
            mouseHook = new ThreadMouseDoubleClickHook(TryHandleMouseDoubleClick);
        }
        catch (Exception ex)
        {
            DialogHelper.ShowWarning("Failed to initialize func_SMILES2IMG:\n" + ex.Message);
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        try
        {
            mouseHook?.Dispose();
            mouseHook = null;

            if (pptEvents != null)
            {
                pptEvents.WindowBeforeDoubleClick -= PowerPointStructureManager.OnWindowBeforeDoubleClick;
                pptEvents = null;
            }
            pptApp = null;
        }
        catch { }
    }

    private bool TryHandleMouseDoubleClick()
    {
        if (isEditing || pptApp == null) return false;
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
                        if (StructureMetadata.IsMolstar(targetShape.AlternativeText, targetShape.Title))
                        {
                            PowerPointStructureManager.EditSelectedMolstar(pptApp, targetShape);
                        }
                        else
                        {
                            PowerPointStructureManager.EditSelectedMolecule(pptApp, targetShape);
                        }
                    }
                    finally
                    {
                        isEditing = false;
                    }
                    return true; // Suppress PowerPoint's default double-click action (e.g. Picture Format pane)
                }
            }
        }
        catch { }
        return false;
    }

    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public bool ExecuteRibbonAction(string actionName)
    {
        var method = GetType().GetMethod(actionName, BindingFlags.Public | BindingFlags.Instance);
        if (method == null) return false;
        method.Invoke(this, new object?[] { null });
        return true;
    }

    public string GetCustomUI(string RibbonID)
    {
        return $@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabSmiles"" label=""SMILES"">
        <group id=""grpInsertMolecules"" label=""Insert Structure"">
          <button id=""btnInsertMolecule""
                  label=""Insert / Edit Structure""
                  screentip=""EPAM Ketcher Molecule Editor""
                  supertip=""Open Ketcher web editor to draw chemical structures and insert into the current slide.""
                  imageMso=""PensGallery""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertMolstar""
                  label=""3D Structure (Mol*)""
                  screentip=""Mol* 3D Molecular Viewer""
                  supertip=""Open Mol* 3D Web Viewer to visualize macromolecules, load PDB structures, and insert 2D pictures or interactive 3D models into the active slide.""
                  imageMso=""Insert3DModelFromFile""
                  size=""large""
                  onAction=""OnInsertMolstar"" />
          <button id=""btnQuickInsert""
                  label=""Paste Structure""
                  screentip=""Paste Structure from Clipboard""
                  supertip=""Convert SMILES or KET directly into a molecular structure image without opening the editor.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
        </group>
        <group id=""grpStructureTools"" label=""Structure Tools"">
          <button id=""btnCopySmiles""
                  label=""Copy SMILES""
                  screentip=""Copy SMILES of Selected Molecule""
                  supertip=""Extract SMILES string from the selected molecule and copy it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyInchi""
                  label=""Copy InChI""
                  screentip=""Copy IUPAC InChI""
                  supertip=""Extract or convert the selected molecular structure into an IUPAC InChI string and copy to clipboard.""
                  imageMso=""SignatureLineInsert""
                  size=""large""
                  onAction=""OnCopyInchi"" />
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <button id=""btnPptCopySvg""
                  label=""Copy SVG""
                  screentip=""Copy Vector Graphic (SVG)""
                  supertip=""Renders the selected molecule as a scalable vector graphic (SVG) and copies it to the clipboard for Illustrator or Inkscape.""
                  imageMso=""CopyPicture""
                  size=""large""
                  onAction=""OnCopySvg"" />
{RibbonXmlBuilder.BuildRibbonMoleculeInfoMenu("Ppt")}
{RibbonXmlBuilder.BuildRibbonBatchStyleMenu(
    "menuRibbonPptBatchStyle",
    "Batch Style Normalizer",
    "Convert all molecule images in the active presentation to a unified publication style (Color, B&amp;W, White BG, or Dark Mode).",
    "btnPptBatchColorTrans",
    "btnPptBatchBw",
    "btnPptBatchWhiteBg",
    "btnPptBatchDark")}
        </group>
        <group id=""grpPptAbout"" label=""About"">
          <button id=""btnPptAbout""
                  label=""About SMILES""
                  screentip=""About func_SMILES2IMG""
                  supertip=""Displays creator info, GitHub repository link, and open-source third-party licenses.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnAbout"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuPicture",
    "menuContextPptPicture",
    "btnContextPptPicture",
    "sepContextPptPicture",
    "menuContextPptPictureMoleculeInfo",
    "PptPicture",
    "sepContextPptPicture")}
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuShape",
    "menuContextPptShape",
    "btnContextPptShape",
    "sepContextPptShape",
    "menuContextPptShapeMoleculeInfo",
    "PptShape",
    "sepContextPptShape")}
    <contextMenu idMso=""ContextMenuFrame"">
      <menu id=""menuContextPptFrame""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextPptFrameInsert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextPptFrameInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextPptFrameQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
      </menu>
      <menuSeparator id=""sepContextPptFrame"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
{RibbonXmlBuilder.BuildTextOrTableContextMenu(
    "ContextMenuText",
    "menuContextPptText",
    "btnContextPptText",
    "sepContextPptText",
    "menuContextPptTextMoleculeInfo",
    "PptText",
    "sepContextPptText")}
{RibbonXmlBuilder.BuildTextOrTableContextMenu(
    "ContextMenuTable",
    "menuContextPptTable",
    "btnContextPptTable",
    "sepContextPptTable",
    "menuContextPptTableMoleculeInfo",
    "PptTable",
    "sepContextPptTable")}
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

    public void OnCopyInchi(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeInchi(pptApp);
    }

    public void OnCopyKet(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeKet(pptApp);
    }

    public void OnCopySvg(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeSvg(pptApp);
    }

    public void OnBatchStyleColorTrans(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.ApplyBatchStyle(pptApp, "transparent", "color", "Standard Color (Transparent)");
    }

    public void OnBatchStyleBw(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.ApplyBatchStyle(pptApp, "transparent", "bw", "Publication B&W");
    }

    public void OnBatchStyleWhiteBg(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.ApplyBatchStyle(pptApp, "white", "color", "White Background (Color)");
    }

    public void OnBatchStyleDark(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.ApplyBatchStyle(pptApp, "transparent", "white", "Dark Mode (White Ink)");
    }

    public void OnCopyMoleculeInfo(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.CopySelectedMoleculeInfo(pptApp);
    }

    public void OnCopySubstanceName(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedSubstanceName(pptApp);
    }

    public void OnCopyFormula(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedFormula(pptApp);
    }

    public void OnCopyMolecularWeight(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedMolecularWeight(pptApp);
    }

    public void OnCopyInchiKey(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedInchiKey(pptApp);
    }

    public void OnCopyAllSummary(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedAllSummary(pptApp);
    }

    public void OnCopyMultiLine(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedMultiLine(pptApp);
    }

    public void OnCopyTsvRow(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedTsvRow(pptApp);
    }

    public void OnCopyMarkdown(IRibbonControl control)
    {
        if (pptApp != null) PowerPointStructureManager.CopySelectedMarkdown(pptApp);
    }

    public void OnInsertMolstar(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.InsertMolstarStructure(pptApp);
    }

    public void OnEditMolstar(IRibbonControl control)
    {
        if (pptApp == null) return;
        PowerPointStructureManager.EditSelectedMolstar(pptApp);
    }

    public void OnAbout(IRibbonControl control)
    {
        IntPtr hwnd = IntPtr.Zero;
        try
        {
            if (pptApp != null)
                hwnd = (IntPtr)pptApp.HWND;
        }
        catch { }
        DialogHelper.ShowAbout(hwnd);
    }
}
