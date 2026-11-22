# Word 리본 · 우클릭 메뉴

Microsoft Word 상단 **`[SMILES]`** 리본 탭 전체 명령표와 5가지 영역별 **`Chemical Structure`** 우클릭 컨텍스트 메뉴 레퍼런스입니다.

<FeatureMedia id="word-ribbon-tools" />

## 1. 상단 `[SMILES]` 리본 전체 명령표

| 리본 그룹 | 버튼 / 드롭다운 메뉴 | Word에서의 동작 |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit Structure` | Ketcher 2D를 열어 커서 위치에 새 구조식을 삽입하거나(텍스트 선택 시 자동 프리로드) 선택한 2D 그림을 재편집 |
| | `3D Structure (Mol*)` | Mol\* 3D를 열어 `2D Picture`(`.png`) 또는 회전 `3D Model (GLB)` 삽입(PDB 코드 선택 시 자동 로드) 및 재편집 |
| | `Paste Structure` | 드래그 선택한 SMILES·InChI·KET·화합물명/CAS 텍스트를 `0.1초` 만에 인라인 구조식으로 치환 (선택 없으면 클립보드 변환) |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` | 선택한 구조식 그림(또는 드래그한 텍스트)에서 표준 SMILES, IUPAC InChI, 무손실 KET JSON 복사 |
| | `Copy SVG` | TOC Graphic, Adobe Illustrator, Inkscape용 무손실 벡터 SVG 그래픽 + XML 텍스트 복사 |
| | `Molecule Info` | 개별 물성, 17열 요약표, TSV 행, SI용 줄바꿈 목록, Markdown 복사 및 17종 상세 물성 창 열기 |
| | `Batch Style` | 드래그 선택한 영역(선택이 없으면 **Word 문서 전체**)의 2D 구조식을 4종 프리셋으로 원클릭 일괄 변환 (`Mol*` 제외) |
| **About** | `Online Docs` | 웹 브라우저에서 Word 공식 문서 카테고리 인트로(`/word/`) 즉시 열기 |
| | `About func_SMILES2IMG` | 설치 버전, 내장 엔진 버전, 서드파티 라이선스 확인 및 원클릭 최신 업데이트 체크 |

## 2. 우클릭 컨텍스트 메뉴 (5가지 영역 지원)

Word의 **인라인 그림(`InlinePicture`)**, **부동 그림(`Picture`)**, **도형/3D 개체(`Shape`)**, **본문 텍스트(`Text`)**, **표 셀 내부(`Table`)** 어디서 우클릭하든 3단계 **`Chemical Structure`** 메뉴가 표시됩니다:

![Microsoft Word 3단계 Chemical Structure 우클릭 메뉴](/word-context-menu-guide.png)

## 관련 가이드

- [**더블클릭 · 재편집 문제 해결**](/ko/word/troubleshooting/double-click) — 그림 서식 창이 열릴 때 Alt Text 태그 점검 및 공동저자 공유 안내
- [**표 잘림 · 일괄 스타일 문제 해결**](/ko/word/troubleshooting/table-layout) — 표 셀 상단 그림 잘림, 줄 간격 `고정값` 해제, `Batch Style` 범위 안내
- [**1. 2D 구조식 · 반응식 작성**](/ko/word/structures) · [**2. 화합물 표(SAR) & 실험 물성(SI)**](/ko/word/tables-properties) — 입력 방식 및 표 레이아웃 가이드
- [**리본 탭 미표시 해결**](/ko/guide/troubleshooting/ribbon-tab) — Word COM 추가 기능 격리 해제 방법
