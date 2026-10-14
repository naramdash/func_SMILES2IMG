using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using Smiles2Img.COM;
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

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            wordApp = (MSWord.Application)Application;
            var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            NativeIndigo.Initialize(asmDir);

            wordEvents = (MSWord.ApplicationEvents4_Event)wordApp;
            wordEvents.WindowBeforeDoubleClick += WordStructureManager.OnWindowBeforeDoubleClick;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to initialize SMILES for Word:\n" + ex.Message, "SMILES2IMG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        try
        {
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
                  supertip=""Open Ketcher web editor to draw chemical structures and insert into the document.""
                  imageMso=""EquationInsertNew""
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
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyInchi"" />
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <menu id=""menuRibbonWordMoleculeInfo""
                label=""Molecule Info""
                screentip=""Copy Molecule Info &amp; Substance Properties""
                supertip=""Directly copy substance name, formula, MW, InChI, or open the detailed molecule properties window.""
                imageMso=""GroupTaskProperties""
                size=""large"">
            <button id=""btnRibbonWordCopyName""
                    label=""Copy Substance Name""
                    imageMso=""PropertySheet""
                    onAction=""OnCopySubstanceName"" />
            <button id=""btnRibbonWordCopyFormula""
                    label=""Copy Chemical Formula""
                    imageMso=""EquationInsertNew""
                    onAction=""OnCopyFormula"" />
            <button id=""btnRibbonWordCopyMW""
                    label=""Copy Molecular Weight (MW)""
                    imageMso=""Calculate""
                    onAction=""OnCopyMolecularWeight"" />
            <button id=""btnRibbonWordSubCopySmiles""
                    label=""Copy SMILES""
                    imageMso=""Copy""
                    onAction=""OnCopySmiles"" />
            <button id=""btnRibbonWordSubCopyInchi""
                    label=""Copy InChI""
                    imageMso=""Copy""
                    onAction=""OnCopyInchi"" />
            <button id=""btnRibbonWordSubCopyInchiKey""
                    label=""Copy InChIKey""
                    imageMso=""Copy""
                    onAction=""OnCopyInchiKey"" />
            <button id=""btnRibbonWordCopyAllSummary""
                    label=""Copy All Summary (Text)""
                    imageMso=""FileDocumentInspect""
                    onAction=""OnCopyAllSummary"" />
            <menuSeparator id=""sepRibbonWordSubDialog"" />
            <button id=""btnRibbonWordShowDialog""
                    label=""Detailed Molecule Info Window...""
                    imageMso=""GroupTaskProperties""
                    onAction=""OnCopyMoleculeInfo"" />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuInlinePicture"">
      <button id=""btnContextInlineEdit""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineEditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyInchi""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menu id=""menuContextInlineMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextInlineCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextInlineCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextInlineCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculate""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextInlineSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextInlineSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextInlineSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""Copy""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextInlineCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <menuSeparator id=""sepContextInlineSubDialog"" />
        <button id=""btnContextInlineShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextInline"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEdit""
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
      <menu id=""menuContextPictureMoleculeInfo""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""
            insertBeforeMso=""Cut"">
        <button id=""btnContextPictureCopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContextPictureCopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContextPictureCopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculate""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContextPictureSubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextPictureSubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextPictureSubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""Copy""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContextPictureCopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <menuSeparator id=""sepContextPictureSubDialog"" />
        <button id=""btnContextPictureShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>
      <menuSeparator id=""sepContextPicture"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEdit""
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
      <menu id=""menuContextShapeMoleculeInfo""
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
                imageMso=""Calculate""
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
      <menuSeparator id=""sepContextShape"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuText"">
      <menu id=""menuContextWordText""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextTextInsert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextTextInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextTextQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextWordTextSubCopy"" />
        <button id=""btnContextTextCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextTextCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextTextCopyKet""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <menuSeparator id=""sepContextWordTextSubInfo"" />
        <menu id=""menuContextWordTextMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextTextCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextTextCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextTextCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculate""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextTextSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextTextSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextTextSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""Copy""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextTextCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <menuSeparator id=""sepContextTextSubDialog"" />
          <button id=""btnContextTextShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextText"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuTable"">
      <menu id=""menuContextWordTable""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""btnContextTableInsert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""btnContextTableInsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""btnContextTableQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""sepContextWordTableSubCopy"" />
        <button id=""btnContextTableCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContextTableCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContextTableCopyKet""
                label=""Copy Canvas Data (KET)""
                imageMso=""Copy""
                onAction=""OnCopyKet"" />
        <menuSeparator id=""sepContextWordTableSubInfo"" />
        <menu id=""menuContextWordTableMoleculeInfo""
              label=""Copy Molecule Info""
              imageMso=""GroupTaskProperties"">
          <button id=""btnContextTableCopyName""
                  label=""Copy Substance Name""
                  imageMso=""PropertySheet""
                  onAction=""OnCopySubstanceName"" />
          <button id=""btnContextTableCopyFormula""
                  label=""Copy Chemical Formula""
                  imageMso=""EquationInsertNew""
                  onAction=""OnCopyFormula"" />
          <button id=""btnContextTableCopyMW""
                  label=""Copy Molecular Weight (MW)""
                  imageMso=""Calculate""
                  onAction=""OnCopyMolecularWeight"" />
          <button id=""btnContextTableSubCopySmiles""
                  label=""Copy SMILES""
                  imageMso=""Copy""
                  onAction=""OnCopySmiles"" />
          <button id=""btnContextTableSubCopyInchi""
                  label=""Copy InChI""
                  imageMso=""Copy""
                  onAction=""OnCopyInchi"" />
          <button id=""btnContextTableSubCopyInchiKey""
                  label=""Copy InChIKey""
                  imageMso=""Copy""
                  onAction=""OnCopyInchiKey"" />
          <button id=""btnContextTableCopyAllSummary""
                  label=""Copy All Summary (Text)""
                  imageMso=""FileDocumentInspect""
                  onAction=""OnCopyAllSummary"" />
          <menuSeparator id=""sepContextTableSubDialog"" />
          <button id=""btnContextTableShowDialog""
                  label=""Detailed Molecule Info Window...""
                  imageMso=""GroupTaskProperties""
                  onAction=""OnCopyMoleculeInfo"" />
        </menu>
      </menu>
      <menuSeparator id=""sepContextTable"" insertAfterMso=""PasteGalleryMini"" />
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
}
