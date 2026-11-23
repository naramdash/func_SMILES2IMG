# 자주 묻는 질문 (FAQ)

스타일 변환, 벡터 SVG 복사, `.sdf` 파일 연동, 물성 표 복사, 업데이트 등 자주 묻는 활용 질문 모음입니다. 공통 오류는 [설치 보안 차단](/ko/guide/troubleshooting/install), [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab), [문서 공유](/ko/guide/troubleshooting/document-sharing)에서, 앱별 오류는 [Excel](/ko/excel/troubleshooting/formula-errors), [Word](/ko/word/troubleshooting/double-click), [PowerPoint](/ko/powerpoint/troubleshooting/dark-slides), [Editor](/ko/editor/troubleshooting/webview2-loading) 트러블슈팅에서 확인하세요.

## 1. 논문용 스타일 · 벡터(SVG)

### Q. 논문 투고용 흑백(ACS 스타일), 다크 모드, 또는 원하는 선 색상으로 바꿀 수 있나요?
* **문서 전체 일괄 변환:** Excel, Word, PowerPoint 상단 리본에서 **`[SMILES]` ➔ `[Batch Style]`** 메뉴를 클릭하면 선택한 영역이나 문서 전체의 구조식을 **논문용 흑백**(`Publication B&W`), **다크 모드**(`Dark Mode`), **표준 컬러**(`Standard Color`), **흰색 배경**(`White Background`)으로 통일합니다.
* **엑셀 수식에서 개별 지정 (커스텀 선 색상 포함):**
  - 흑백 선화: `=SMILES2IMG(A2, "", "bw")`
  - 다크 모드 (투명 배경 + 흰색 선): `=SMILES2IMG(A2, "trans", "white")`
  - 원하는 브랜드/테마 선 색상 지정: `=SMILES2IMG(A2, "", "ink=#1e40af")`

### Q. Adobe Illustrator나 Inkscape용 벡터(`.svg`)로 복사할 수 있나요?
네! 분자 셀이나 그림을 선택하고 상단 리본 **`[SMILES]` ➔ `[Copy SVG]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**)를 클릭하면 확대해도 깨지지 않는 벡터 그래픽으로 복사되어 바로 붙여넣을 수 있습니다.

## 2. `.sdf` 파일 & 분자 물성 표 복사

### Q. 엑셀에서 `.sdf` 화학 데이터 파일을 열 수 있나요?
네! 엑셀 상단 **`[SMILES]` ➔ `[Import SDF]`** 버튼을 사용하세요:
* **`Import SDF (.sdf)...`:** `.sdf` 파일을 읽어 새 시트에 11개 물성·식별자 열(`Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `TPSA`, `Rings`, `Aromatic Rings`, `Fsp3`)을 갖춘 라이브 수식 표를 자동 생성합니다.

### Q. 분자량, 화학식, LogP, TPSA 정보를 표로 한 번에 복사하려면?
분자 셀이나 그림을 **우클릭 ➔ `Chemical Structure` ➔ `Copy Molecule Info`** 메뉴에서 원하는 서식을 고르세요:
* **`Copy All Summary (Text)`:** **머리글(Header) + 데이터** 2줄 표로 복사되어 새 표를 만들 때 적합합니다.
* **`Copy as Excel Row (TSV Columns)`:** **데이터 행만** 12개 칸으로 복사되어 기존 표 아래에 이어 붙일 때 씁니다.
* **`Copy as Multi-Line List`:** PowerPoint 슬라이드 텍스트 상자에 글머리 기호로 붙여넣기 좋습니다.
* **`Copy as Markdown Table` / JSON:** Notion, GitHub, 연구노트용 (`Detailed Molecule Info Window...` 하단 드롭다운에서 서식 선택 후 **`[Copy All]`** 또는 개별 행 더블클릭 복사 가능).

## 3. 보안 · 호환성 · 업데이트

### Q. 그린 구조식이 외부 서버로 전송되나요?
**아니요. 핵심 2D 렌더링, 2D 편집, 모든 물성·SAR 계산은 내 PC에서 로컬로 처리됩니다.** 인터넷은 미캐시 PubChem 이름/CAS 조회, Mol\* PDB 코드 다운로드, 업데이트 확인을 직접 호출할 때만 선택적으로 사용됩니다. 자세한 내용은 [로컬 우선 보안 & 문서 공유](/ko/guide/#privacy)를 참고하세요.

### Q. `=SMILES2IMG`와 `=SMILES2IMG.FLOAT` 중 무엇을 써야 하나요?
* **Microsoft 365 / Excel 2024 이상:** 셀 안 네이티브 이미지로 정렬·필터와 연동되는 `=SMILES2IMG(A2)`를 사용하세요.
* **Excel 2016 / 2019 / 2021:** 뒤에 `.FLOAT`를 붙인 `=SMILES2IMG.FLOAT(A2)`를 사용하면 셀 크기에 맞춰 움직이는 그림 도형으로 생성됩니다.

### Q. 새 버전 업데이트는 어떻게 하나요?
상단 리본 **`[SMILES]` ➔ `[About func_SMILES2IMG]`** 창에서 **`[Check for Updates]`** 버튼을 누르거나, [PowerShell 설치 명령](/ko/guide/install)을 다시 실행하면 최신 버전으로 업그레이드됩니다:

<FeatureMedia id="common-about-dialog" />

## 관련 가이드

- [**설치 가이드**](/ko/guide/install) — PowerShell 1줄 자동 설치, 오프라인 기업망 수동 설치 및 업데이트/제거 방법
- [**49개 엑셀 함수 전체 개요**](/ko/excel/formulas) — 전체 49종 함수 규격, 반환 데이터 형식 및 공용 스타일 키워드 사전
