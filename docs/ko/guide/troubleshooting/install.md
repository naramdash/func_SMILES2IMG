# 설치 보안 차단

Windows 11 스마트 앱 컨트롤 경고, 수동 다운로드 `.xll` / `.zip` 차단, 사내 방화벽 환경 문제를 해결합니다.

## 1. Windows 11 스마트 앱 컨트롤(SAC) 경고

설치 직후 Windows 11에서 **"파일 형식과 확장명이 일치하지 않습니다"** 경고가 뜨거나 엑셀에서 `#NAME?` 오류가 표시되는 경우:

* **원인:** Windows 11 **스마트 앱 컨트롤**이 강제 켜짐 상태일 때 서명되지 않은 오픈소스 `.xll` 추가 기능 실행을 차단합니다.
* **안심하세요:** 스마트 앱 컨트롤을 꺼도 **Microsoft Defender 백신**(실시간 바이러스 감시)은 100% 그대로 작동합니다.
* **10초 해결법:**
  1. Windows **설정** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **개인 정보 및 보안** ➔ **Windows 보안**
  2. **앱 및 브라우저 컨트롤** ➔ **스마트 앱 컨트롤 설정** 클릭
  3. **꺼짐(Off)** 상태로 변경한 뒤 오피스를 **완전히 닫고 다시 실행**합니다.

## 2. 수동 다운로드한 `.xll` 또는 `.zip` 파일 차단

웹 브라우저로 직접 내려받은 파일에는 Windows 보안 차단 플래그(Mark-of-the-Web)가 붙습니다.

* **수동 해결:** 압축 풀기 전 `.zip` 또는 `.xll` 파일을 우클릭 ➔ **속성** ➔ 하단 **[차단 해제]**(Unblock) 체크 ➔ **[확인]** 클릭.
* **자동 대안:** [10초 PowerShell 자동 설치](/ko/guide/install)를 실행하면 설치된 모든 파일에 `Unblock-File`이 자동 적용됩니다.

## 3. 사내 보안망 · 오프라인 환경

* **GitHub 직접 접속이 막힌 경우:** [오프라인 수동 ZIP 설치](/ko/guide/install#offline-zip) 절차를 따르세요.
* **오프라인 완전 지원 기능:** 핵심 2D 렌더링(`=SMILES2IMG`, `=KET2IMG`, `=INCHI2IMG`), 모든 물성·Lipinski·SAR 계산, 내장 용매/의약품/CAS 사전, Ketcher 2D 편집, `.sdf` 입출력, Batch Style, SVG 복사는 인터넷 없이 100% 로컬로 작동합니다.
* **인터넷 필요 기능:** 내장 사전에 없는 PubChem 이름/CAS 조회, Mol\* 4자리 PDB 코드 다운로드, **`[Check for Updates]`** 버튼.

## 관련 트러블슈팅

- [**리본 탭 미표시**](/ko/guide/troubleshooting/ribbon-tab) — Excel, Word, PowerPoint에서 숨겨지거나 비활성화된 [SMILES] 탭 복구
- [**엑셀 수식 오류**](/ko/excel/troubleshooting/formula-errors) — #NAME?, #VALUE!, #N/A 및 .FLOAT 호환성 문제 해결
