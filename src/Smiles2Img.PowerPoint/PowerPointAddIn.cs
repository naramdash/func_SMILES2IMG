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
        <group id=""grpInsertMolecules"" label=""Insert Structure"">
          <button id=""btnInsertMolecule""
                  label=""Insert / Edit Structure""
                  screentip=""EPAM Ketcher Molecule Editor""
                  supertip=""Open Ketcher web editor to draw chemical structures and insert into the current slide.""
                  imageMso=""DrawingCanvasInsert""
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
                  imageMso=""SourceCodeInsert""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <button id=""btnPptCopySvg""
                  label=""Copy SVG""
                  screentip=""Copy Vector Graphic (SVG)""
                  supertip=""Renders the selected molecule as a scalable vector graphic (SVG) and copies it to the clipboard for Illustrator or Inkscape.""
                  imageMso=""CopyPicture""
                  size=""large""
                  onAction=""OnCopySvg"" />
          <menu id=""menuRibbonPptMoleculeInfo""
                label=""Molecule Info""
                screentip=""Copy Molecule Info &amp; Substance Properties""
                supertip=""Directly copy substance name, formula, MW, InChI, or open the detailed molecule properties window.""
                imageMso=""FileProperties""
                size=""large"">
            <button id=""btnRibbonPptCopyName""
                    label=""Copy Substance Name""
                    imageMso=""FontName""
                    onAction=""OnCopySubstanceName"" />
            <button id=""btnRibbonPptCopyFormula""
                    label=""Copy Chemical Formula""
                    imageMso=""Subscript""
                    onAction=""OnCopyFormula"" />
            <button id=""btnRibbonPptCopyMW""
                    label=""Copy Molecular Weight (MW)""
                    imageMso=""Calculator""
                    onAction=""OnCopyMolecularWeight"" />
            <button id=""btnRibbonPptSubCopySmiles""
                    label=""Copy SMILES""
                    imageMso=""Copy""
                    onAction=""OnCopySmiles"" />
            <button id=""btnRibbonPptSubCopyInchi""
                    label=""Copy InChI""
                    imageMso=""SignatureLineInsert""
                    onAction=""OnCopyInchi"" />
            <button id=""btnRibbonPptSubCopyInchiKey""
                    label=""Copy InChIKey""
                    imageMso=""Key""
                    onAction=""OnCopyInchiKey"" />
            <button id=""btnRibbonPptCopyAllSummary""
                    label=""Copy All Summary (Text)""
                    imageMso=""DocumentSummary""
                    onAction=""OnCopyAllSummary"" />
            <button id=""btnRibbonPptCopyMultiLine""
                    label=""Copy as Multi-Line List""
                    imageMso=""Bullets""
                    onAction=""OnCopyMultiLine"" />
            <button id=""btnRibbonPptCopyTsvRow""
                    label=""Copy as Excel Row (TSV Columns)""
                    imageMso=""TableInsert""
                    onAction=""OnCopyTsvRow"" />
            <button id=""btnRibbonPptCopyMarkdown""
                    label=""Copy as Markdown Table""
                    imageMso=""TableCustom""
                    onAction=""OnCopyMarkdown"" />
            <menuSeparator id=""sepRibbonPptSubDialog"" />
            <button id=""btnRibbonPptShowDialog""
                    label=""Detailed Molecule Info Window...""
                    imageMso=""FileProperties""
                    onAction=""OnCopyMoleculeInfo"" />
          </menu>
          <menu id=""menuRibbonPptBatchStyle""
                label=""Batch Style""
                screentip=""Batch Style Normalizer""
                supertip=""Convert all molecule images in the active presentation to a unified publication style (Color, B&amp;W, White BG, or Dark Mode).""
                imageMso=""PictureRecolorGallery""
                size=""large"">
            <button id=""btnPptBatchColorTrans""
                    label=""Standard Color (Transparent)""
                    imageMso=""ColorMenu""
                    onAction=""OnBatchStyleColorTrans"" />
            <button id=""btnPptBatchBw""
                    label=""Publication B&amp;W (ACS Monochrome)""
                    imageMso=""BlackAndWhiteGrayscale""
                    onAction=""OnBatchStyleBw"" />
            <button id=""btnPptBatchWhiteBg""
                    label=""White Background (Color)""
                    imageMso=""ShapeFillColorPicker""
                    onAction=""OnBatchStyleWhiteBg"" />
            <button id=""btnPptBatchDark""
                    label=""Dark Mode (White Ink)""
                    imageMso=""HighContrast""
                    onAction=""OnBatchStyleDark"" />
          </menu>
        </group>
        <group id=""grpPptAbout"" label=""About"">
          <button id=""btnPptAbout""
                  label=""About SMILES""
                  screentip=""About SMILES for Office""
                  supertip=""Displays creator info, GitHub repository link, and open-source third-party licenses.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnAbout"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPptPictureEdit""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureEditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopyInchi""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptPictureCopySvg""
              label=""Copy Vector Graphic (SVG)""
              imageMso=""CopyPicture""
              onAction=""OnCopySvg""
              insertBeforeMso=""Cut"" />
      <menu id=""menuContextPptPictureMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextPptPictureCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextPptPictureCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextPptPictureCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculator""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextPptPictureSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPptPictureSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPptPictureSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""SignatureLineInsert""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextPptPictureCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <button id=""btnContextPptPictureCopyMultiLine""
                label=""Copy as Multi-Line List""
                imageMso=""Bullets""
                onAction=""OnCopyMultiLine"" />
        <menuSeparator id=""sepContextPptPictureSubDialog"" />
        <button id=""btnContextPptPictureShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextPptPicture"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextPptShapeEdit""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeEditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopyInchi""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPptShapeCopySvg""
              label=""Copy Vector Graphic (SVG)""
              imageMso=""CopyPicture""
              onAction=""OnCopySvg""
              insertBeforeMso=""Cut"" />
      <menu id=""menuContextPptShapeMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextPptShapeCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextPptShapeCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextPptShapeCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculator""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextPptShapeSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPptShapeSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPptShapeSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""SignatureLineInsert""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextPptShapeCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <button id=""btnContextPptShapeCopyMultiLine""
                label=""Copy as Multi-Line List""
                imageMso=""Bullets""
                onAction=""OnCopyMultiLine"" />
        <menuSeparator id=""sepContextPptShapeSubDialog"" />
        <button id=""btnContextPptShapeShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextPptShape"" insertBeforeMso=""Cut"" />
    </contextMenu>
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
    <contextMenu idMso=""ContextMenuText"">
      <menu id=""menuContextPptText""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextPptTextInsert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextPptTextInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextPptTextQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextPptTextSubCopy"" />
        <button id=""btnContextPptTextCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPptTextCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPptTextCopyKet""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <button id=""btnContextPptTextCopySvg""
                label=""Copy Vector Graphic (SVG)""
                imageMso=""CopyPicture""
                onAction=""OnCopySvg"" />
        <menuSeparator id=""sepContextPptTextSubInfo"" />
        <menu id=""menuContextPptTextMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextPptTextCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextPptTextCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextPptTextCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculator""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextPptTextSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextPptTextSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextPptTextSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""SignatureLineInsert""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextPptTextCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <button id=""btnContextPptTextCopyMultiLine""
                  label=""Copy as Multi-Line List""
                  imageMso=""Bullets""
                  onAction=""OnCopyMultiLine"" />
          <menuSeparator id=""sepContextPptTextSubDialog"" />
          <button id=""btnContextPptTextShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextPptText"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuTable"">
      <menu id=""menuContextPptTable""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextPptTableInsert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextPptTableInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextPptTableQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextPptTableSubCopy"" />
        <button id=""btnContextPptTableCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPptTableCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPptTableCopyKet""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <button id=""btnContextPptTableCopySvg""
                label=""Copy Vector Graphic (SVG)""
                imageMso=""CopyPicture""
                onAction=""OnCopySvg"" />
        <menuSeparator id=""sepContextPptTableSubInfo"" />
        <menu id=""menuContextPptTableMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextPptTableCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextPptTableCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextPptTableCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculator""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextPptTableSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextPptTableSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextPptTableSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""SignatureLineInsert""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextPptTableCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <button id=""btnContextPptTableCopyMultiLine""
                  label=""Copy as Multi-Line List""
                  imageMso=""Bullets""
                  onAction=""OnCopyMultiLine"" />
          <menuSeparator id=""sepContextPptTableSubDialog"" />
          <button id=""btnContextPptTableShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextPptTable"" insertAfterMso=""PasteGalleryMini"" />
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
        try
        {
            IntPtr hwnd = IntPtr.Zero;
            if (pptApp != null)
            {
                try { hwnd = (IntPtr)pptApp.HWND; } catch { }
            }
            IWin32Window? owner = hwnd != IntPtr.Zero ? new Win32Window(hwnd) : null;
            Smiles2Img.Editor.AboutDialog.ShowModal(owner);
        }
        catch
        {
            Smiles2Img.Editor.AboutDialog.ShowModal();
        }
    }
}
