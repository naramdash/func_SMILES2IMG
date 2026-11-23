# Word 개요

Microsoft Word 안에서 2D 화학 구조식, 다단계 합성 반응식, 3D 단백질 Figure를 직접 그리고, 표 열 너비를 유지한 채 더블클릭으로 재편집하거나 학술지 흑백 스타일로 일괄 변환할 수 있습니다.

<FeatureMedia id="word-overview" />

## 주요 기능

### 1. 표 열 너비를 유지하는 더블클릭 제자리 수정
화합물 요약 표나 본문에 넣은 구조식에서 작용기나 촉매를 고쳐야 할 때 그림을 **더블클릭**하면 Ketcher 2D가 다시 열립니다. 수정 후 적용하면 가로세로 비율과 기존 셀 너비·위치가 그대로 유지됩니다.

<FeatureMedia id="word-ch3-compound-table" />

### 2. 문서 전체 ACS 흑백 일괄 통일 & 벡터 `SVG` 복사
컬러 초안으로 작성한 본문·표 구조식을 투고 전 **`[Batch Style]`** 로 흑백 선화로 일괄 변환할 수 있습니다. TOC Graphic이나 멀티패널 Figure용으로는 Illustrator에 붙여넣을 수 있는 벡터 `SVG` 복사를 지원합니다.

<FeatureMedia id="word-ch6-publication-finish" />

### 3. 계산된 화학식·HRMS 정밀질량 복사 & 3D 단백질 Figure (`Mol*`)
구조식 그림이나 선택한 텍스트에서 화학식·분자량·HRMS 정밀질량(`Da`)·LogP·TPSA를 복사해 Experimental Section에 바로 붙여넣을 수 있으며, 4자리 PDB 코드(`1CRN`, `6LU7`)로 3D 단백질 Figure도 삽입할 수 있습니다.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## 작업 방식 비교

| Word 논문 · 보고서 작업 | 기존 방식 (ChemDraw / 화면 캡처) | `func_SMILES2IMG` 사용 시 |
| :--- | :--- | :--- |
| **화합물 표 안 구조식 수정** | 외부 툴 재오픈 → PNG 재저장 → 삭제·재삽입 → 열 너비 재조정 | **구조식 더블클릭 → Ketcher 수정 → <kbd>Ctrl</kbd> + <kbd>Enter</kbd>** |
| **컬러 초안 일괄 흑백 변환** | 파일 수십 개를 하나씩 흑백으로 다시 저장해 교체 | **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`** |
| **Experimental 화학식 · HRMS 기재** | 별도 계산·검색 후 직접 타이핑 | **`[Molecule Info]`로 12종 물성·정밀질량 복사** |
| **Mac 공동저자 · 심사위원 `.docx` 공유** | 전용 OLE 개체 깨짐 또는 Mac에서 편집 오류 | **표준 고해상도 PNG + `.glb` 저장 — Mac·웹·PDF 호환** |
| **결합 포켓 3D 단백질 Figure 제작** | PyMOL/ChimeraX 실행 → 캡처·여백 크롭 → 재캡처 반복 | **`[3D Structure (Mol*)]`로 삽입 & 더블클릭 시 카메라 각도 유지** |

## 로컬 실행 및 호환성

<OfficeOverviewCta app="word" />

## 설치 후 기본 체험

1. **텍스트 구조식 변환:** 빈 Word 문단에 `Aspirin` 또는 `CC(=O)Oc1ccccc1C(=O)O`를 입력해 드래그 선택한 뒤 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 눌러 보세요.
2. **더블클릭 제자리 수정:** 들어간 구조식을 **더블클릭**해 **Ketcher 2D**에서 치환기를 고치고 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 눌러 제자리 교체를 확인하세요(새로 그릴 때는 **`[Insert / Edit Structure]`** 클릭).
3. **자체 완결형 Playbook(`.docx`) 실습:** <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Word_Playbook.docx" download="func_SMILES2IMG_Word_Playbook.docx"><strong>func_SMILES2IMG_Word_Playbook.docx</strong></a>를 내려받아 SAR 표, KET 반응식, Mol\* `1CRN` 그림, ACS 흑백 변환을 순서대로 따라 해 보세요 ([6페이지 상세 구성 →](/ko/word/playbook)).

## 작업별 가이드 바로가기

| Word에서 하려는 작업 | 가이드 | 핵심 명령 |
| :--- | :--- | :--- |
| 논문 투고 · 특허 SAR 표 · SI · 3D 보고서 · 학위논문 시나리오 보기 | [Word 활용 사례](/ko/word/use-cases) | 5가지 R&D · 학위논문 시나리오 |
| 본문 구조식 그리기, SMILES·화합물명 변환, 무손실 `KET` 반응식 | [1. 2D 구조식 · 반응식 작성](/ko/word/structures) | `[Insert / Edit Structure]` / `[Paste Structure]` |
| 열 너비 고정·줄 간격 잘림 방지 SAR 표 제작 & HRMS/SI 물성 복사 | [2. 화합물 표(SAR) & 실험 물성(SI)](/ko/word/tables-properties) | 열 너비 고정 + `[Molecule Info]` |
| 4자리 PDB 코드·로컬 `.pdb`로 논문용 3D 단백질 Figure(`Mol*`) 삽입 | [3. 3D 단백질 Figure (Mol*)](/ko/word/3d-figures) | `[3D Structure (Mol*)]` |
| 투고 전 체크리스트(ACS 흑백 통일, TOC 벡터 SVG, SI 추출, Mac 공유) | [4. 투고 마감 · 벡터 SVG · 공동저자 공유](/ko/word/publishing) | `[Batch Style]` / `[Copy SVG]` |
| 상단 `[SMILES]` 리본 전체 기능 및 5종 우클릭 메뉴 살펴보기 | [리본 · 우클릭 메뉴](/ko/word/reference) | 상단 `[SMILES]` 탭 & 우클릭 |
