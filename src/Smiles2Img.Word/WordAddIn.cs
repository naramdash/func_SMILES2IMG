using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using Smiles2Img.COM;
using Smiles2Img.Editor;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

[ComVisible(true)]
[Guid("41A5C1F0-559A-46E0-B9C0-8B35A0F1D910")]
[ProgId("Smiles2Img.Word.AddIn")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class WordAddIn : IDTExtensibility2, IRibbonExtensibility
{
    private MSWord.Application? wordApp;
    private MSWord.ApplicationEvents4_Event? wordEvents;
    private IRibbonUI? ribbon;
    private ThreadMouseDoubleClickHook? mouseHook;

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            wordApp = (MSWord.Application)Application;
            try
            {
                if (AddInInst != null)
                    ((dynamic)AddInInst).Object = this;
            }
            catch { }
            var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            NativeIndigo.Initialize(asmDir);

            wordEvents = (MSWord.ApplicationEvents4_Event)wordApp;
            wordEvents.WindowBeforeDoubleClick += WordStructureManager.OnWindowBeforeDoubleClick;
            mouseHook = new ThreadMouseDoubleClickHook(() =>
                wordApp != null && WordStructureManager.TryHandleFloatingDoubleClick(wordApp));
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

            if (wordEvents != null)
            {
                wordEvents.WindowBeforeDoubleClick -= WordStructureManager.OnWindowBeforeDoubleClick;
                wordEvents = null;
            }
            wordApp = null;
        }
        catch { }
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
                  supertip=""Open Ketcher web editor to draw chemical structures and insert into the document.""
                  imageMso=""PensGallery""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertMolstar""
                  label=""3D Structure (Mol*)""
                  screentip=""Mol* 3D Macromolecular Viewer""
                  supertip=""Open Mol* 3D viewer to render macromolecular structures (PDB/AlphaFold/CIF/SDF) as interactive 3D models or publication images.""
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
          <button id=""btnWordCopySvg""
                  label=""Copy SVG""
                  screentip=""Copy Scalable Vector Graphics (SVG)""
                  supertip=""Renders the selected molecule or reaction into crisp vector SVG and copies it to the clipboard for Office 365, Illustrator, Inkscape, or HTML.""
                  imageMso=""CopyPicture""
                  size=""large""
                  onAction=""OnCopySvg"" />
{RibbonXmlBuilder.BuildRibbonMoleculeInfoMenu("Word")}
{RibbonXmlBuilder.BuildRibbonBatchStyleMenu(
    "menuRibbonWordBatchStyle",
    "Batch Normalize Document Molecule Styles",
    "Convert all molecular structures in the active Word document to a unified style preset in one click.",
    "btnWordBatchStyleColorTrans",
    "btnWordBatchStyleBw",
    "btnWordBatchStyleWhiteBg",
    "btnWordBatchStyleDark")}
        </group>
        <group id=""grpWordAbout"" label=""About"">
          <button id=""btnWordAbout""
                  label=""About func_SMILES2IMG""
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
    "ContextMenuInlinePicture",
    "menuContextWordInline",
    "btnContextInline",
    "sepContextWordInline",
    "menuContextInlineMoleculeInfo",
    "Inline",
    "sepContextInline")}
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuPicture",
    "menuContextWordPicture",
    "btnContextPicture",
    "sepContextWordPicture",
    "menuContextPictureMoleculeInfo",
    "Picture",
    "sepContextPicture")}
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuShape",
    "menuContextWordShape",
    "btnContextShape",
    "sepContextWordShape",
    "menuContextShapeMoleculeInfo",
    "Shape",
    "sepContextShape")}
{RibbonXmlBuilder.BuildTextOrTableContextMenu(
    "ContextMenuText",
    "menuContextWordText",
    "btnContextText",
    "sepContextWordText",
    "menuContextWordTextMoleculeInfo",
    "Text",
    "sepContextText")}
{RibbonXmlBuilder.BuildTextOrTableContextMenu(
    "ContextMenuTable",
    "menuContextWordTable",
    "btnContextTable",
    "sepContextWordTable",
    "menuContextWordTableMoleculeInfo",
    "Table",
    "sepContextTable")}
  </contextMenus>
</customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        ribbon = ribbonUI;
    }

    public void OnInsertMolecule(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.InsertMolecule(wordApp);
    }

    public void OnInsertMolstar(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.InsertMolstarStructure(wordApp);
    }

    public void OnQuickInsert(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.QuickInsertFromSelectionOrClipboard(wordApp);
    }

    public void OnEditMolecule(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.EditSelectedMolecule(wordApp);
    }

    public void OnEditMolstar(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.EditSelectedMolstar(wordApp);
    }

    public void OnCopySmiles(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeSmiles(wordApp);
    }

    public void OnCopyInchi(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeInchi(wordApp);
    }

    public void OnCopyKet(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeKet(wordApp);
    }

    public void OnCopySvg(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeSvg(wordApp);
    }

    public void OnBatchStyleColorTrans(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.ApplyBatchStyle(wordApp, "trans", "color", "Standard Color (Transparent)");
    }

    public void OnBatchStyleBw(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.ApplyBatchStyle(wordApp, "trans", "bw", "Publication B&W");
    }

    public void OnBatchStyleWhiteBg(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.ApplyBatchStyle(wordApp, "white", "color", "White Background");
    }

    public void OnBatchStyleDark(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.ApplyBatchStyle(wordApp, "trans", "white", "Dark Mode (White Ink)");
    }

    public void OnCopyMoleculeInfo(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeInfo(wordApp);
    }

    public void OnCopySubstanceName(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedSubstanceName(wordApp);
    }

    public void OnCopyFormula(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedFormula(wordApp);
    }

    public void OnCopyMolecularWeight(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedMolecularWeight(wordApp);
    }

    public void OnCopyInchiKey(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedInchiKey(wordApp);
    }

    public void OnCopyAllSummary(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedAllSummary(wordApp);
    }

    public void OnCopyMultiLine(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedMultiLine(wordApp);
    }

    public void OnCopyTsvRow(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedTsvRow(wordApp);
    }

    public void OnCopyMarkdown(IRibbonControl control)
    {
        if (wordApp != null) WordStructureManager.CopySelectedMarkdown(wordApp);
    }

    public void OnAbout(IRibbonControl control)
    {
        IntPtr hwnd = IntPtr.Zero;
        try
        {
            if (wordApp?.ActiveWindow != null)
                hwnd = (IntPtr)wordApp.ActiveWindow.Hwnd;
        }
        catch { }
        DialogHelper.ShowAbout(hwnd);
    }
}
