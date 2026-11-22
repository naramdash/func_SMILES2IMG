using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using Smiles2Img.Editor;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

[ComVisible(true)]
public class SmilesRibbon : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
    {
        return $@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabSmiles"" label=""SMILES"">
        <group id=""grpInsertMolecules"" label=""Insert Structure"">
          <button id=""btnInsertMolecule""
                  label=""Insert / Edit (SMILES)""
                  screentip=""Insert / Edit Structure (SMILES2IMG)""
                  supertip=""Opens Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=SMILES2IMG).""
                  imageMso=""PensGallery""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertInchiMolecule""
                  label=""Insert / Edit (InChI)""
                  screentip=""Insert / Edit Structure (INCHI2IMG)""
                  supertip=""Opens Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=INCHI2IMG).""
                  imageMso=""TextBoxInsert""
                  size=""large""
                  onAction=""OnInsertInchiMolecule"" />
          <button id=""btnInsertKetMolecule""
                  label=""Insert / Edit (KET)""
                  screentip=""Insert / Edit Structure (KET2IMG)""
                  supertip=""Opens Ketcher editor to sketch chemical structures, reactions, and canvases and insert them into the active cell as formulas (=KET2IMG).""
                  imageMso=""SmartArtInsert""
                  size=""large""
                  onAction=""OnInsertKetMolecule"" />
          <button id=""btnInsertMolstar""
                  label=""3D Structure (Mol*)""
                  screentip=""Mol* 3D Macromolecular Viewer""
                  supertip=""Opens Mol* 3D viewer to download PDB structures, visualize macromolecules, and insert interactive 3D models (.glb) or 2D images onto the worksheet.""
                  imageMso=""Insert3DModelFromFile""
                  size=""large""
                  onAction=""OnInsertMolstar"" />
          <button id=""btnQuickInsert""
                  label=""Paste Structure""
                  screentip=""Paste Structure from Clipboard or Cell""
                  supertip=""Instantly converts clipboard or active-cell SMILES, InChI, KET, or chemical names/CAS (e.g. Aspirin) into a structure formula without opening the editor.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
        </group>
        <group id=""grpStructureTools"" label=""Structure Tools"">
          <button id=""btnCopySmiles""
                  label=""Copy SMILES""
                  screentip=""Copy Canonical SMILES""
                  supertip=""Extracts and copies canonical SMILES from the selected formula cell or molecule image to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyInchi""
                  label=""Copy InChI""
                  screentip=""Copy IUPAC InChI""
                  supertip=""Extracts and copies IUPAC InChI string from the selected formula cell or molecule image to the clipboard.""
                  imageMso=""SignatureLineInsert""
                  size=""large""
                  onAction=""OnCopyInchi"" />
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected formula cell or molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <button id=""btnCopySvg""
                  label=""Copy SVG""
                  screentip=""Copy Scalable Vector Graphics (SVG)""
                  supertip=""Renders the selected molecule or reaction into crisp vector SVG and copies it to the clipboard for Office 365, Illustrator, Inkscape, or HTML.""
                  imageMso=""CopyPicture""
                  size=""large""
                  onAction=""OnCopySvg"" />
{RibbonXmlBuilder.BuildRibbonMoleculeInfoMenu("Excel")}
{RibbonXmlBuilder.BuildRibbonBatchStyleMenu(
    "menuRibbonExcelBatchStyle",
    "Batch Normalize Molecule Styles",
    "Convert all molecular structures on the active worksheet to a unified style preset in one click.",
    "btnBatchStyleColorTrans",
    "btnBatchStyleBw",
    "btnBatchStyleWhiteBg",
    "btnBatchStyleDark")}
          <menu id=""menuRibbonExcelSdf""
                label=""SDF File""
                screentip=""Import or Export SDF Chemical Files""
                supertip=""Import an MDL .sdf chemical table into the worksheet or export worksheet structures to an .sdf file.""
                imageMso=""DatabaseQueryNew""
                size=""large"">
            <button id=""btnImportSdf""
                    label=""Import SDF File (.sdf)...""
                    imageMso=""ImportTextFile""
                    onAction=""OnImportSdf"" />
            <button id=""btnExportSdf""
                    label=""Export Sheet to SDF (.sdf)...""
                    imageMso=""FileSaveAs""
                    onAction=""OnExportSdf"" />
          </menu>
        </group>
        <group id=""grpDiagnostics"" label=""Diagnostics"">
          <button id=""btnOpenDiagnostics""
                  label=""Diagnostics""
                  screentip=""SMILES2IMG Diagnostics &amp; Log""
                  supertip=""Displays parser logs and add-in diagnostic information.""
                  imageMso=""WatchWindow""
                  size=""large""
                  onAction=""OnShowDiagnostics"" />
        </group>
{RibbonXmlBuilder.BuildRibbonAboutGroup("grpAbout", "btnOpenDocs", "btnAbout", "Excel")}
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
{RibbonXmlBuilder.BuildExcelRangeContextMenu(
    "ContextMenuCell",
    "sepContextCellPre",
    "menuContextExcelCell",
    "btnContext",
    "btnContextInsertMolstarCell",
    "btnContextQuickInsertCell",
    "sepContextCell",
    "btnContext",
    "menuContextExcelCellMoleculeInfo",
    "Cell",
    "sepContextMolecule")}
{RibbonXmlBuilder.BuildExcelRangeContextMenu(
    "ContextMenuListRange",
    "sepContextListPre",
    "menuContextExcelList",
    "btnContextList",
    "btnContextListInsertMolstar",
    "btnContextListQuickInsertCell",
    "sepContextList",
    "btnContextList",
    "menuContextExcelListMoleculeInfo",
    "List",
    "sepContextListMolecule")}
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuPicture",
    "menuContextExcelPicture",
    "btnContextPicture",
    "sepContextExcelPicture",
    "menuContextPictureExcelMoleculeInfo",
    "Pic",
    "sepContextPicMolecule")}
{RibbonXmlBuilder.BuildObjectContextMenu(
    "ContextMenuShape",
    "menuContextExcelShape",
    "btnContextShape",
    "sepContextExcelShape",
    "menuContextShapeExcelMoleculeInfo",
    "Shape",
    "sepContextShapeMolecule")}
  </contextMenus>
</customUI>";
    }

    private static void RunMacroSafe(Action action)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SmilesRibbon] Macro execution error: {ex}");
            }
        });
    }

    private static void RunDirectSafe(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SmilesRibbon] Direct UI execution error: {ex}");
        }
    }

    public void OnInsertMolecule(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.OpenKetcherCommand());

    public void OnInsertInchiMolecule(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.OpenKetcherInchiCommand());

    public void OnInsertKetMolecule(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.OpenKetcherKetCommand());

    public void OnInsertCanvasMolecule(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.InsertCanvasMoleculeCommand());

    public void OnQuickInsert(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.QuickInsertCommand());

    public void OnEditMolecule(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.EditSelectedMoleculeCommand());

    public void OnCopySmiles(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMoleculeSmilesCommand());

    public void OnCopyInchi(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMoleculeInchiCommand());

    public void OnCopyKet(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMoleculeKetCommand());

    public void OnCopySvg(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMoleculeSvgCommand());

    public void OnBatchStyleColorTrans(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.BatchStyleColorTransCommand());

    public void OnBatchStyleBw(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.BatchStyleBwCommand());

    public void OnBatchStyleWhiteBg(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.BatchStyleWhiteBgCommand());

    public void OnBatchStyleDark(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.BatchStyleDarkCommand());

    public void OnImportSdf(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.ImportSdfCommand());

    public void OnExportSdf(IRibbonControl control)
        => RunMacroSafe(() => KetcherEditor.ExportSdfCommand());

    public void OnCopyMoleculeInfo(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMoleculeInfoCommand());

    public void OnCopySubstanceName(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedSubstanceNameCommand());

    public void OnCopyFormula(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedFormulaCommand());

    public void OnCopyMolecularWeight(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMolecularWeightCommand());

    public void OnCopyInchiKey(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedInchiKeyCommand());

    public void OnCopyAllSummary(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedAllSummaryCommand());

    public void OnCopyTsvRow(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedTsvRowCommand());

    public void OnCopyMultiLine(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMultiLineCommand());

    public void OnCopyMarkdown(IRibbonControl control)
        => RunDirectSafe(() => KetcherEditor.CopySelectedMarkdownCommand());

    public void OnInsertMolstar(IRibbonControl control)
    {
        RunMacroSafe(() =>
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.InsertMolstarStructure(app);
        });
    }

    public void OnEditMolstar(IRibbonControl control)
    {
        RunMacroSafe(() =>
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.EditSelectedMolstar(app);
        });
    }

    public void OnShowDiagnostics(IRibbonControl control)
        => RunMacroSafe(() => LogDisplay.Show());

    public void OnOpenDocs(IRibbonControl control)
        => RunDirectSafe(() => DialogHelper.OpenUrl(AboutDialog.GetDocsUrl("Excel")));

    public void OnAbout(IRibbonControl control)
        => RunMacroSafe(() => DialogHelper.ShowAbout(ExcelDnaUtil.WindowHandle, "Excel"));
}
