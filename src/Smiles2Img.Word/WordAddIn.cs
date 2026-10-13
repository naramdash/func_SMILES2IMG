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
            MessageBox.Show("SMILES for Word 초기화 실패:\n" + ex.Message, "SMILES2IMG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        <group id=""grpMolecules"" label=""분자 구조식 (Structure)"">
          <button id=""btnInsertMolecule""
                  label=""분자 구조식 삽입""
                  screentip=""EPAM Ketcher 분자 에디터""
                  supertip=""Ketcher 웹 에디터를 열어 화학 구조식을 그리고 현재 커서 위치에 삽입합니다.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnQuickInsert""
                  label=""클립보드 즉시 삽입""
                  screentip=""클립보드 또는 선택된 SMILES/KET 즉시 삽입""
                  supertip=""에디터를 열지 않고 클립보드나 문서 내 선택된 SMILES/KET 문자열을 즉시 분자 구조식 이미지로 변환하여 삽입합니다.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
          <button id=""btnEditMolecule""
                  label=""구조식 편집""
                  screentip=""선택된 분자 구조식 편집""
                  supertip=""선택된 분자 이미지를 Ketcher 에디터로 다시 불러와 수정합니다.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
          <button id=""btnCopySmiles""
                  label=""SMILES 복사""
                  screentip=""선택된 분자 구조식의 SMILES 복사""
                  supertip=""선택된 분자 이미지에서 SMILES 문자열을 추출하여 클립보드에 복사합니다.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""분자 정보 복사""
                  screentip=""화학식 및 분자량(MW) 정보 복사""
                  supertip=""선택된 분자 구조식의 화학식(Formula), 분자량(MW), 정밀 질량(Exact Mass)을 계산하여 표시하고 클립보드에 복사합니다.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuInlinePicture"">
      <button id=""btnContextInlineEdit""
              label=""분자 구조식 편집 (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopySmiles""
              label=""SMILES 복사 (추출)""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyInfo""
              label=""분자 정보 (화학식·분자량) 복사""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInlineCopyKet""
              label=""캔버스 데이터(KET) 복사""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextInline"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEdit""
              label=""분자 구조식 편집 (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopySmiles""
              label=""SMILES 복사 (추출)""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyInfo""
              label=""분자 정보 (화학식·분자량) 복사""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextPictureCopyKet""
              label=""캔버스 데이터(KET) 복사""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextPicture"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEdit""
              label=""분자 구조식 편집 (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopySmiles""
              label=""SMILES 복사 (추출)""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyInfo""
              label=""분자 정보 (화학식·분자량) 복사""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextShapeCopyKet""
              label=""캔버스 데이터(KET) 복사""
              imageMso=""SaveObjectAs""
              onAction=""OnCopyKet""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextShape"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuText"">
      <button id=""btnContextTextQuickInsert""
              label=""선택/클립보드 SMILES → 구조식 삽입""
              imageMso=""EquationInsertNew""
              onAction=""OnQuickInsert""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextText"" insertBeforeMso=""Cut"" />
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
