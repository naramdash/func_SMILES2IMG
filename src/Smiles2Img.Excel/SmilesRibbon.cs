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
        <group id=""grpMolecules"" label=""Molecular Structures"">
          <button id=""btnInsertMolecule""
                  label=""Insert Structure""
                  screentip=""EPAM Ketcher Molecular Editor (Ctrl+Shift+K)""
                  supertip=""Opens the EPAM Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=SMILES2IMG or =KET2IMG).""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertCanvasMolecule""
                  label=""Insert Canvas (KET)""
                  screentip=""Insert Freehand Canvas Structure (KET)""
                  supertip=""Opens the EPAM Ketcher editor to draw reactions, arrows, and multi-molecule canvases as high-resolution sheet images.""
                  imageMso=""SmartArtInsert""
                  size=""large""
                  onAction=""OnInsertCanvasMolecule"" />
          <button id=""btnQuickInsert""
                  label=""Quick Paste""
                  screentip=""Quick Paste SMILES/KET from Clipboard""
                  supertip=""Instantly converts clipboard SMILES or KET data into a structure image without opening the editor.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
          <button id=""btnEditMolecule""
                  label=""Edit Structure""
                  screentip=""Edit Selected Structure""
                  supertip=""Reloads the selected formula cell (=SMILES2IMG, =KET2IMG) or sheet canvas image back into Ketcher for editing.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
          <button id=""btnCopySmiles""
                  label=""Copy SMILES""
                  screentip=""Copy Canonical SMILES""
                  supertip=""Extracts and copies canonical SMILES from the selected formula cell or molecule image to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""Copy Molecule Info""
                  screentip=""Calculate &amp; Copy Formula, MW, Exact Mass""
                  supertip=""Calculates chemical formula, molecular weight (MW), and exact mass, displays properties, and copies them to the clipboard.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
          <button id=""btnOpenDiagnostics""
                  label=""Diagnostics""
                  screentip=""SMILES2IMG Diagnostics &amp; Log""
                  supertip=""Displays parser logs and add-in diagnostic information.""
                  imageMso=""MacroSecurity""
                  size=""normal""
                  onAction=""OnShowDiagnostics"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuCell"">
      <button id=""btnContextEditMoleculeCell""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextCopySmilesCell""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextCopyInfoCell""
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInsertMolecule""
              label=""Insert Structure (Formula)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInsertCanvasMolecule""
              label=""Insert Canvas Structure (KET)...""
              imageMso=""SmartArtInsert""
              onAction=""OnInsertCanvasMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuListRange"">
      <button id=""btnContextListEditMoleculeCell""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListCopySmilesCell""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListCopyInfoCell""
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListInsertMolecule""
              label=""Insert Structure (Formula)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListInsertCanvasMolecule""
              label=""Insert Canvas Structure (KET)...""
              imageMso=""SmartArtInsert""
              onAction=""OnInsertCanvasMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextListMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyInfo""
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPicMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEditMolecule""
              label=""Edit Structure (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyInfo""
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
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
