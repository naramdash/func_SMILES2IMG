# Word 개요

치환기 하나 고칠 때마다 외부 드로잉 툴을 다시 열고, 캡처하고, 깨진 표 열 너비를 맞추는 반복 작업을 멈추세요. Microsoft Word 안에서 2D 화학 구조식·다단계 합성 반응식·3D 단백질 Figure를 직접 그리고, 더블클릭으로 제자리 수정하고, 클릭 한 번으로 학술지 흑백 마감까지 끝낼 수 있습니다.

<FeatureMedia id="word-overview" />

## 논문 · 보고서 작성에서 달라지는 3가지 결정적 순간

### 1. 표 열 너비 깨짐 없이 더블클릭 0.05초 제자리 수정
화합물 요약 표나 본문에 넣은 구조식에서 작용기나 촉매를 고쳐야 할 때 그림을 **더블클릭**만 하세요. 가로세로 비율이 고정된 채 기존 셀 폭과 위치 그대로 즉시 교체됩니다.

<FeatureMedia id="word-ch3-compound-table" />

### 2. 투고 전날 밤, 클릭 한 번으로 문서 전체 ACS 흑백 통일 & 벡터 `SVG`
컬러 초안으로 그려 둔 수십 개의 본문·표 구조식을 투고 직전에 클릭 한 번으로 순수 흑백 선화로 바꿉니다. TOC Graphic이나 멀티패널 Figure가 필요할 때는 Illustrator로 바로 붙여넣을 수 있는 벡터 `SVG`로 복사됩니다.

<FeatureMedia id="word-ch6-publication-finish" />

### 3. 오타 없는 화학식·HRMS 정밀질량 복사 & 3D 단백질 Figure (`Mol*`)
구조식 그림(또는 드래그한 텍스트)에서 화학식·분자량·HRMS 정밀질량(`Da`)·LogP·TPSA를 즉시 복사해 Experimental Section을 채우고, 4자리 PDB 코드(`1CRN`, `6LU7`)로 논문용 3D 단백질 그림까지 바로 삽입합니다.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## 기존 방식 vs. `func_SMILES2IMG`

| Word 논문 · 보고서 작업 | 기존 방식 (ChemDraw / 화면 캡처) | `func_SMILES2IMG` 사용 시 |
| :--- | :--- | :--- |
| **화합물 표 안 구조식 오타 수정** | 외부 툴 재오픈 → PNG 재저장 → 삭제·재삽입 → 열 너비 재조정 | **구조식 더블클릭 → Ketcher 수정 → <kbd>Ctrl</kbd> + <kbd>Enter</kbd>** |
| **투고 전 컬러 초안 30개 흑백 변환** | 파일 수십 개를 하나씩 흑백으로 다시 저장해 교체 | **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]` 1클릭** |
| **Experimental 화학식 · HRMS 기재** | 별도 계산·검색 후 직접 타이핑(오타·반올림 실수 위험) | **`[Molecule Info]`로 12종 물성·정밀질량 1클릭 복사** |
| **Mac 공동저자 · 심사위원 `.docx` 공유** | 전용 OLE 개체 깨짐 또는 Mac에서 편집/렌더링 오류 | **표준 고해상도 PNG + `.glb` 저장 — Mac·웹·PDF 100% 호환** |
| **결합 포켓 3D 단백질 Figure 제작** | PyMOL/ChimeraX 실행 → 캡처·여백 크롭 → 각도 바꿀 때마다 재캡처 | **`[3D Structure (Mol*)]`로 삽입 & 더블클릭 시 카메라 각도 그대로 재오픈** |

## 연구실에서 안심하고 바로 설치할 수 있는 이유

<OfficeOverviewCta app="word" />

## 설치 후 1분 안에 체험하기

1. **0.1초 즉시 변환 체험:** 빈 Word 문단에 `Aspirin` 또는 `CC(=O)Oc1ccccc1C(=O)O`를 입력해 드래그 선택한 뒤 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 눌러 보세요.
2. **더블클릭 제자리 수정:** 방금 들어간 구조식을 **더블클릭**해 **Ketcher 2D**에서 치환기를 고치고 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 눌러 제자리 교체를 확인하세요(새로 그릴 때는 **`[Insert / Edit Structure]`** 클릭).
3. **자체 완결형 Playbook(`.docx`) 실습:** <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Word_Playbook.docx" download="func_SMILES2IMG_Word_Playbook.docx"><strong>func_SMILES2IMG_Word_Playbook.docx</strong></a>를 내려받아 SAR 표, KET 반응식, Mol\* `1CRN` 그림, 1클릭 ACS 흑백 변환을 5분 만에 순서대로 따라 해 보세요 ([6페이지 상세 구성 →](/ko/word/playbook)).

## 작업별 실전 가이드 바로가기

| Word에서 하려는 작업 | 가이드 | 핵심 명령 |
| :--- | :--- | :--- |
| 논문 투고 · 특허 SAR 표 · SI · 3D 보고서 · 학위논문 실전 시나리오 보기 | [Word 활용 사례](/ko/word/use-cases) | 5가지 R&D · 학위논문 실전 시나리오 |
| 본문 구조식 그리기, SMILES·화합물명 `0.1초` 변환, 무손실 `KET` 반응식 | [1. 2D 구조식 · 반응식 작성](/ko/word/structures) | `[Insert / Edit Structure]` / `[Paste Structure]` |
| 열 너비 고정·줄 간격 잘림 방지 SAR 표 제작 & 오타 없는 HRMS/SI 물성 복사 | [2. 화합물 표(SAR) & 실험 물성(SI)](/ko/word/tables-properties) | 열 너비 고정 + `[Molecule Info]` |
| 4자리 PDB 코드·로컬 `.pdb`로 논문용 3D 단백질 Figure(`Mol*`) 삽입 | [3. 3D 단백질 Figure (Mol*)](/ko/word/3d-figures) | `[3D Structure (Mol*)]` |
| 투고 전 5분 체크리스트(ACS 흑백 통일, TOC 벡터 SVG, SI 추출, Mac 공유) | [4. 투고 마감 · 벡터 SVG · 공동저자 공유](/ko/word/publishing) | `[Batch Style]` / `[Copy SVG]` |
| 상단 `[SMILES]` 리본 전체 기능 및 5종 우클릭 메뉴 살펴보기 | [리본 · 우클릭 메뉴](/ko/word/reference) | 상단 `[SMILES]` 탭 & 우클릭 |
