using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;

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
          <button id=""btnInsertKetMolecule""
                  label=""Insert / Edit (KET)""
                  screentip=""Insert / Edit Structure (KET2IMG)""
                  supertip=""Opens Ketcher editor to sketch chemical structures, reactions, and canvases and insert them into the active cell as formulas (=KET2IMG).""
                  imageMso=""SmartArtInsert""
                  size=""large""
                  onAction=""OnInsertKetMolecule"" />
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
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected formula cell or molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""Molecule Info""
                  screentip=""Molecule Info &amp; Substance Properties""
                  supertip=""Display and copy chemical substance name, IUPAC name, molecular weight (MW), formula, and SMILES.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
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
        <button id=""btnContextInsertKetMolecule""
                label=""Insert / Edit Structure (KET2IMG)...""
                imageMso=""SmartArtInsert""
                onAction=""OnInsertKetMolecule"" />
        <button id=""btnContextQuickInsertCell""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextCellSubCopy"" />
        <button id=""btnContextCopySmilesCell""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextCopyKetCell""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <button id=""btnContextCopyInfoCell""
                label=""Molecule Info (Name, MW &amp; Formula)...""
                imageMso=""Copy""
                onAction=""OnCopyMoleculeInfo"" />
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
        <button id=""btnContextListInsertKetMolecule""
                label=""Insert / Edit Structure (KET2IMG)...""
                imageMso=""SmartArtInsert""
                onAction=""OnInsertKetMolecule"" />
        <button id=""btnContextListQuickInsertCell""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextListSubCopy"" />
        <button id=""btnContextListCopySmilesCell""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextListCopyKetCell""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <button id=""btnContextListCopyInfoCell""
                label=""Molecule Info (Name, MW &amp; Formula)...""
                imageMso=""Copy""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextListMolecule"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyInfo""
              label=""Molecule Info (Name, MW &amp; Formula)...""
              imageMso=""Copy""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPicMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyInfo""
              label=""Molecule Info (Name, MW &amp; Formula)...""
              imageMso=""Copy""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
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

    public void OnShowDiagnostics(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            LogDisplay.Show();
        });
    }
}
