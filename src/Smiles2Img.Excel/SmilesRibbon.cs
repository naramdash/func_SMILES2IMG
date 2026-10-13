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
        <group id=""grpMolecules"" label=""분자 구조식 (Structure)"">
          <button id=""btnInsertMolecule""
                  label=""분자 구조식 삽입""
                  screentip=""EPAM Ketcher 분자 에디터 (Ctrl+Shift+K)""
                  supertip=""Ketcher 웹 에디터를 열어 화학 구조식을 그리고 현재 활성 셀에 수식(=SMILES2IMG)으로 삽입합니다.""
                  imageMso=""EquationInsertNew""
                  size=""large""
                  onAction=""OnInsertMolecule"" />
          <button id=""btnInsertCanvasMolecule""
                  label=""캔버스 구조식 삽입""
                  screentip=""캔버스 형태 분자 구조식 삽입 (KET)""
                  supertip=""Ketcher 웹 에디터를 열어 반응식, 화살표, 텍스트가 포함된 분자 구조식을 시트 위에 캔버스 이미지(KET)로 삽입합니다.""
                  imageMso=""SmartArtInsert""
                  size=""large""
                  onAction=""OnInsertCanvasMolecule"" />
          <button id=""btnQuickInsert""
                  label=""클립보드 즉시 삽입""
                  screentip=""클립보드 SMILES/KET 즉시 삽입""
                  supertip=""에디터를 열지 않고 클립보드에 복사된 SMILES 또는 KET 구조식을 현재 셀 또는 시트에 즉시 삽입합니다.""
                  imageMso=""Paste""
                  size=""large""
                  onAction=""OnQuickInsert"" />
          <button id=""btnEditMolecule""
                  label=""구조식 편집""
                  screentip=""선택된 분자 구조식 편집""
                  supertip=""선택된 수식 셀(=SMILES2IMG) 또는 시트 위의 분자 캔버스 이미지를 Ketcher 에디터로 다시 불러와 수정합니다.""
                  imageMso=""PictureEditProperties""
                  size=""large""
                  onAction=""OnEditMolecule"" />
          <button id=""btnCopySmiles""
                  label=""SMILES 복사""
                  screentip=""선택된 셀 또는 이미지의 SMILES 복사""
                  supertip=""선택된 수식 셀 또는 분자 이미지에서 SMILES 문자열을 추출하여 클립보드에 복사합니다.""
                  imageMso=""Copy""
                  size=""large""
                  onAction=""OnCopySmiles"" />
          <button id=""btnCopyMoleculeInfo""
                  label=""분자 정보 복사""
                  screentip=""화학식 및 분자량(MW) 정보 복사""
                  supertip=""선택된 셀 또는 분자 구조식의 화학식(Formula), 분자량(MW), 정밀 질량(Exact Mass)을 계산하여 표시하고 클립보드에 복사합니다.""
                  imageMso=""Info""
                  size=""large""
                  onAction=""OnCopyMoleculeInfo"" />
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
      <button id=""btnContextEditMoleculeCell""
              label=""분자 구조식 편집 (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextCopySmilesCell""
              label=""SMILES 복사 (추출)""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextCopyInfoCell""
              label=""분자 정보 (화학식·분자량) 복사""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInsertMolecule""
              label=""분자 구조식 삽입 (셀 수식)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextInsertCanvasMolecule""
              label=""캔버스 구조식 삽입 (KET)...""
              imageMso=""SmartArtInsert""
              onAction=""OnInsertCanvasMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuListRange"">
      <button id=""btnContextListEditMoleculeCell""
              label=""분자 구조식 편집 (Ketcher)...""
              imageMso=""PictureEditProperties""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListCopySmilesCell""
              label=""SMILES 복사 (추출)""
              imageMso=""Copy""
              onAction=""OnCopySmiles""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListCopyInfoCell""
              label=""분자 정보 (화학식·분자량) 복사""
              imageMso=""Info""
              onAction=""OnCopyMoleculeInfo""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListInsertMolecule""
              label=""분자 구조식 삽입 (셀 수식)...""
              imageMso=""EquationInsertNew""
              onAction=""OnInsertMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""btnContextListInsertCanvasMolecule""
              label=""캔버스 구조식 삽입 (KET)...""
              imageMso=""SmartArtInsert""
              onAction=""OnInsertCanvasMolecule""
              insertBeforeMso=""Cut"" />
      <menuSeparator id=""sepContextListMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuPicture"">
      <button id=""btnContextPictureEditMolecule""
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
      <menuSeparator id=""sepContextPicMolecule"" insertBeforeMso=""Cut"" />
    </contextMenu>
    <contextMenu idMso=""ContextMenuShape"">
      <button id=""btnContextShapeEditMolecule""
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
