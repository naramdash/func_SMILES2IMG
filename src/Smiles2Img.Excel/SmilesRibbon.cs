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
      <tab id=""tabSmiles"" label=""SMILES"" insertAfterMso=""TabInsert"">
        <group id=""grpMolecules"" label=""분자 구조식 (Structure)"">
          <button id=""btnInsertMolecule""
                  label=""분자 구조식 삽입""
                  screentip=""EPAM Ketcher 분자 에디터 (Ctrl+Shift+K)""
                  supertip=""Ketcher 웹 에디터를 열어 화학 구조식을 그리고 현재 활성 셀에 삽입합니다.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnOpenDiagnostics""
                  label=""진단 (Diagnostics)""
                  screentip=""SMILES2IMG 로그 및 진단 창""
                  supertip=""추가 기능 상태 및 렌더링 진단 정보를 확인합니다.""
                  imageMso=""MacroSecurity""
                  size=""normal""
                  onAction=""OnShowDiagnostics"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuCell"">
      <button id=""btnContextInsertMolecule""
              label=""분자 구조식 삽입 (Ketcher)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuListRange"">
      <button id=""btnContextListInsertMolecule""
              label=""분자 구조식 삽입 (Ketcher)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextListMolecule"" insertBeforeMso=""Cut"" />
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

    public void OnShowDiagnostics(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            LogDisplay.Show();
        });
    }
}
