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
      <tab id=""tabSmiles"" label=""SMILES"" insertAfterMso=""TabInsert"">
        <group id=""grpMolecules"" label=""분자 구조식 (Structure)"">
          <button id=""btnInsertMolecule""
                  label=""분자 구조식 삽입""
                  screentip=""EPAM Ketcher 분자 에디터""
                  supertip=""Ketcher 웹 에디터를 열어 화학 구조식을 그리고 현재 커서 위치에 삽입합니다.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnEditMolecule""
                  label=""구조식 편집""
                  screentip=""선택된 분자 구조식 편집""
                  supertip=""선택된 분자 이미지를 Ketcher 에디터로 다시 불러와 수정합니다.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
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

    public void OnEditMolecule(IRibbonControl control)
    {
        if (wordApp == null) return;
        WordStructureManager.EditSelectedMolecule(wordApp);
    }
}
