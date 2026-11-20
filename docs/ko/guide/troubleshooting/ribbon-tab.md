# 리본 탭 미표시

설치 후 Microsoft Excel, Word, PowerPoint 상단에 **`[SMILES]`** 리본 탭이 나타나지 않을 때의 복구 절차입니다.

## 1. 오피스 완전히 종료 후 재시작

추가 기능은 오피스 프로그램이 처음 시작될 때만 로드됩니다.

1. 작업을 저장하고 열려 있는 Excel, Word, PowerPoint 창을 **모두 닫습니다.**
2. <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>Esc</kbd>로 **작업 관리자**를 열어 백그라운드에 `EXCEL.EXE`, `WINWORD.EXE`, `POWERPNT.EXE`가 남아 있지 않은지 확인합니다.
3. 오피스를 다시 실행하고 상단 메뉴에 **`[SMILES]`** 탭이 나타나는지 확인합니다.

## 2. 비활성화된 추가 기능(Disabled Items) 복구

오피스가 비정상 종료된 이력이 있으면 추가 기능을 **사용할 수 없는 항목**으로 자동 격리할 수 있습니다.

1. **파일** ➔ **옵션** ➔ **추가 기능**으로 이동합니다.
2. 하단 **관리** 드롭다운에서 **사용할 수 없는 항목** 선택 ➔ **이동...** 클릭.
3. 목록에 `func_SMILES2IMG` 또는 `Smiles2Img`가 있다면 선택 후 **사용**을 누르고 오피스를 재시작합니다.

## 3. 활성 추가 기능 체크 상태 확인

* **Microsoft Excel:** **파일** ➔ **옵션** ➔ **추가 기능** ➔ **관리: Excel 추가 기능** ➔ **이동...** 버튼을 눌러 `func_SMILES2IMG` (`Smiles2Img.Excel-AddIn64-packed.xll` 등)가 체크되어 있는지 확인합니다.
* **Microsoft Word / PowerPoint:** **파일** ➔ **옵션** ➔ **추가 기능** ➔ **관리: COM 추가 기능** ➔ **이동...** 버튼을 눌러 `func_SMILES2IMG`가 체크되어 있는지 확인합니다.

## 4. 10초 만에 레지스트리 및 파일 복구

[10초 PowerShell 자동 설치](/ko/guide/install) 명령어를 다시 실행하면 관리자 권한 없이 현재 사용자(`HKCU`) 레지스트리와 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 바이너리를 안전하게 복구합니다.

## 관련 트러블슈팅

- [**설치 보안 차단**](/ko/guide/troubleshooting/install) — Windows 11 스마트 앱 컨트롤 및 수동 다운로드 .xll 차단 해제
- [**엑셀 수식 오류**](/ko/excel/troubleshooting/formula-errors) — #NAME?, #VALUE!, .FLOAT 구버전 호환성 해결
