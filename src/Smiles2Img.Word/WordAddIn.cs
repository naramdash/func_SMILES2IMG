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
          <button id=""btnCopyKet""
                  label=""Copy KET""
                  screentip=""Copy Canvas Data (KET)""
                  supertip=""Extracts or converts the selected molecule image into native Ketcher JSON (KET) and copies it to the clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyKet"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""Copy Molecule Info""
                  screentip=""Copy Formula &amp; Molecular Weight""
                  supertip=""Calculate and copy chemical formula, molecular weight (MW), and exact mass to clipboard.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
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
      <button id=""btnContextInlineCopySmiles""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyKet""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyInfo""
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Copy""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextInline"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEdit""
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
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Copy""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPicture"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEdit""
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
              label=""Copy Molecule Info (Formula &amp; MW)""
              imageMso=""Copy""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
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
        <button id=""btnContextTextQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
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
        <button id=""btnContextTableQuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
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

    public void OnCopySmiles(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.CopySelectedMoleculeSmiles(wordApp);
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
}
