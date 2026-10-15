using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

[ComVisible(true)]
public class SmilesRibbon : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
    {
        return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabSmiles"" label=""SMILES"">
        <group id=""grpInsertMolecules"" label=""Insert Structure"">
          <button id=""btnInsertMolecule""
                  label=""Insert / Edit (SMILES)""
                  screentip=""Insert / Edit Structure (SMILES2IMG) (Ctrl+Shift+K)""
                  supertip=""Opens Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=SMILES2IMG).""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertInchiMolecule""
                  label=""Insert / Edit (InChI)""
                  screentip=""Insert / Edit Structure (INCHI2IMG)""
                  supertip=""Opens Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=INCHI2IMG).""
                  imageMso=""EquationInsertNew""
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
                  screentip=""Paste Structure from Clipboard""
                  supertip=""Instantly converts clipboard SMILES or KET data into a structure image without opening the editor.""
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
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyInchi"" />
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected formula cell or molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <menu id=""menuRibbonExcelMoleculeInfo""
                label=""Molecule Info""
                screentip=""Copy Molecule Info &amp; Substance Properties""
                supertip=""Directly copy substance name, formula, MW, InChI, or open the detailed molecule properties window.""
                imageMso=""GroupTaskProperties""
                size=""large"">
            <button id=""btnRibbonExcelCopyName""
                    label=""Copy Substance Name""
                    imageMso=""PropertySheet""
                    onAction=""OnCopySubstanceName"" />
            <button id=""btnRibbonExcelCopyFormula""
                    label=""Copy Chemical Formula""
                    imageMso=""EquationInsertNew""
                    onAction=""OnCopyFormula"" />
            <button id=""btnRibbonExcelCopyMW""
                    label=""Copy Molecular Weight (MW)""
                    imageMso=""Calculator""
                    onAction=""OnCopyMolecularWeight"" />
            <button id=""btnRibbonExcelSubCopySmiles""
                    label=""Copy SMILES""
                    imageMso=""Copy""
                    onAction=""OnCopySmiles"" />
            <button id=""btnRibbonExcelSubCopyInchi""
                    label=""Copy InChI""
                    imageMso=""Copy""
                    onAction=""OnCopyInchi"" />
            <button id=""btnRibbonExcelSubCopyInchiKey""
                    label=""Copy InChIKey""
                    imageMso=""Copy""
                    onAction=""OnCopyInchiKey"" />
            <button id=""btnRibbonExcelCopyAllSummary""
                    label=""Copy All Summary (Text)""
                    imageMso=""FileDocumentInspect""
                    onAction=""OnCopyAllSummary"" />
            <menuSeparator id=""sepRibbonExcelSubDialog"" />
            <button id=""btnRibbonExcelShowDialog""
                    label=""Detailed Molecule Info Window...""
                    imageMso=""GroupTaskProperties""
                    onAction=""OnCopyMoleculeInfo"" />
          </menu>
        </group>
        <group id=""grpDiagnostics"" label=""Diagnostics"">
          <button id=""btnOpenDiagnostics""
                  label=""Diagnostics""
                  screentip=""SMILES2IMG Diagnostics &amp; Log""
                  supertip=""Displays parser logs and add-in diagnostic information.""
                  imageMso=""MacroSecurity""
                  size=""large""
                  onAction=""OnShowDiagnostics"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuCell"">
      <menuSeparator id=""sepContextCellPre"" insertAfterMso=""PasteGalleryMini"" />
      <menu id=""menuContextExcelCell""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextInsertMolecule""
                label=""Insert / Edit Structure (SMILES2IMG)...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextInsertInchiMolecule""
                label=""Insert / Edit Structure (INCHI2IMG)...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertInchiMolecule"" />
        <button id=""btnContextInsertKetMolecule""
                label=""Insert / Edit Structure (KET2IMG)...""
                imageMso=""SmartArtInsert""
                onAction=""OnInsertKetMolecule"" />
        <button id=""btnContextInsertMolstarCell""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextQuickInsertCell""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextCellSubCopy"" />
        <button id=""btnContextCopySmilesCell""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextCopyInchiCell""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextCopyKetCell""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <menuSeparator id=""sepContextCellSubInfo"" />
        <menu id=""menuContextExcelCellMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextCellCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextCellCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextCellCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculator""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextCellSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextCellSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextCellSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""Copy""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextCellCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <menuSeparator id=""sepContextCellSubDialog"" />
          <button id=""btnContextCellShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextMolecule"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuListRange"">
      <menuSeparator id=""sepContextListPre"" insertAfterMso=""PasteGalleryMini"" />
      <menu id=""menuContextExcelList""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextListInsertMolecule""
                label=""Insert / Edit Structure (SMILES2IMG)...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextListInsertInchiMolecule""
                label=""Insert / Edit Structure (INCHI2IMG)...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertInchiMolecule"" />
        <button id=""btnContextListInsertKetMolecule""
                label=""Insert / Edit Structure (KET2IMG)...""
                imageMso=""SmartArtInsert""
                onAction=""OnInsertKetMolecule"" />
        <button id=""btnContextListInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextListQuickInsertCell""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextListSubCopy"" />
        <button id=""btnContextListCopySmilesCell""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextListCopyInchiCell""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextListCopyKetCell""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <menuSeparator id=""sepContextListSubInfo"" />
        <menu id=""menuContextExcelListMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextListCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextListCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextListCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculator""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextListSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextListSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextListSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""Copy""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextListCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <menuSeparator id=""sepContextListSubDialog"" />
          <button id=""btnContextListShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextListMolecule"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureEditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyInchi""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menu id=""menuContextPictureExcelMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextPicCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextPicCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextPicCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculator""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextPicSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPicSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPicSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""Copy""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextPicCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <menuSeparator id=""sepContextPicSubDialog"" />
        <button id=""btnContextPicShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextPicMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeEditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyInchi""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menu id=""menuContextShapeExcelMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextShapeCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextShapeCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextShapeCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculator""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextShapeSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextShapeSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextShapeSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""Copy""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextShapeCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <menuSeparator id=""sepContextShapeSubDialog"" />
        <button id=""btnContextShapeShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextShapeMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
  </contextMenus>
</customUI>";
    }

    public void OnInsertMolecule(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.OpenKetcherCommand();
        });
    }

    public void OnInsertInchiMolecule(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.OpenKetcherInchiCommand();
        });
    }

    public void OnInsertKetMolecule(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.OpenKetcherKetCommand();
        });
    }

    public void OnInsertCanvasMolecule(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.InsertCanvasMoleculeCommand();
        });
    }

    public void OnQuickInsert(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.QuickInsertCommand();
        });
    }

    public void OnEditMolecule(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.EditSelectedMoleculeCommand();
        });
    }

    public void OnCopySmiles(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedMoleculeSmilesCommand();
        });
    }

    public void OnCopyInchi(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedMoleculeInchiCommand();
        });
    }

    public void OnCopyKet(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedMoleculeKetCommand();
        });
    }

    public void OnCopyMoleculeInfo(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedMoleculeInfoCommand();
        });
    }

    public void OnCopySubstanceName(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedSubstanceNameCommand();
        });
    }

    public void OnCopyFormula(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedFormulaCommand();
        });
    }

    public void OnCopyMolecularWeight(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedMolecularWeightCommand();
        });
    }

    public void OnCopyInchiKey(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedInchiKeyCommand();
        });
    }

    public void OnCopyAllSummary(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            KetcherEditor.CopySelectedAllSummaryCommand();
        });
    }

    public void OnInsertMolstar(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.InsertMolstarStructure(app);
        });
    }

    public void OnEditMolstar(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.EditSelectedMolstar(app);
        });
    }

    public void OnShowDiagnostics(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            LogDisplay.Show();
        });
    }
}
