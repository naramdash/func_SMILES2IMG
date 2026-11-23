# 4. 투고 마감 · 벡터 SVG · 공동저자 공유

ACS, RSC, Wiley, Elsevier, Nature 등 학술지나 특허청 제출 직전, Microsoft Word 원고와 PDF 변환본을 완벽하게 마무리하는 **투고 전 5분 체크리스트**입니다.

<FeatureMedia id="word-ch6-publication-finish" />

## 투고 전 5분 마감 체크리스트

### Step 1. 원고 전체 또는 특정 표만 ACS 흑백으로 통일 (`Batch Style`)

컬러 초안으로 작성한 본문과 SAR 표의 구조식을 하나씩 다시 열어 바꿀 필요가 없습니다:

1. **Word 적용 범위 지정:**
   * **원고 전체 (인라인 + 부동 그림 모두):** 아무 그림도 선택하지 않고 본문 아무 곳에나 커서만 둡니다.
   * **특정 표·문단만 변환:** 바꾸고 싶은 화합물 표 행이나 문단만 마우스로 드래그 선택합니다.
2. 상단 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** 를 클릭합니다.
   * 모든 2D 분자 구조식과 `KET` 합성 반응식이 투명 배경 + 순수 흑백 선화로 교체되며, 기존 크기와 열 너비 고정 레이아웃을 그대로 유지합니다(`Mol*` 3D 단백질 Figure는 자동으로 건너뜁니다).
   * 학위논문이나 컬러 프리프린트용으로 되돌릴 때는 **`[Standard Color (Transparent)]`** 를 누르면 즉시 복원됩니다(4종 프리셋 전체 규격은 [스타일 & 출력 상세 매뉴얼](/ko/editor/publishing) 참고).

### Step 2. TOC Graphic · 멀티패널 Figure용 무손실 벡터 `SVG` 추출

**Graphical Abstract(TOC)** 나 `Figure 1A–C` 같은 복합 Figure를 **Adobe Illustrator**나 **Inkscape**에서 정밀 편집해야 할 때:

1. Word 문서의 2D 구조식 또는 KET 반응식 그림을 클릭합니다(또는 본문의 SMILES 텍스트를 드래그 선택).
2. 상단 **`[SMILES]` ➔ `[Copy SVG]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**)를 클릭합니다.
3. Illustrator, Inkscape, 또는 최신 Microsoft 365 문서에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣으면 비트맵이 아닌 편집 가능한 벡터 패스(Path)로 들어갑니다.

### Step 3. Supporting Information(SI) 식별자 · 물성 오타 없이 추출

* **SI 화합물 특성 블록:** 원고 안 구조식 그림을 선택하고 **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** 를 누르면 화학식, 평균 분자량, HRMS 정밀질량(`Da`), LogP, TPSA, SMILES, InChI, InChIKey가 정리된 줄바꿈 목록이 복사됩니다.
* **투고 시스템 입력용 식별자:** 투고 포털에서 SMILES나 InChI 제출을 요구할 때는 그림을 선택하고 **`[Copy SMILES]`** 또는 **`[Copy InChI]`** 를 눌러 즉시 복사하세요.

### Step 4. 공동저자(Mac) · 심사위원 `.docx` 호환성 확인

* **설치 없는 환경에서도 100% 동일 출력:** Word 구조식은 표준 고해상도 PNG 그림과 표준 `.glb` 3D 모델로 문서에 내장됩니다. `func_SMILES2IMG`가 없는 **macOS**, Windows PC, Word 웹, 학술지 PDF 변환기에서도 그림 깨짐 없이 열람·인쇄됩니다.
* **교정본 회신 후에도 재편집 유지:** 화학 메타데이터가 그림 **대체 텍스트(Alt Text)** 와 `CustomXMLParts`에 함께 저장되므로, Mac을 쓰는 공동저자가 본문을 수정해 돌려보내더라도 Windows에서 다시 열면 더블클릭 편집 기능이 그대로 동작합니다.

::: tip 논문 반응식을 PowerPoint 다크모드 발표 자료로 3초 만에 옮기기
논문 발표 슬라이드를 준비 중이라면 Word의 KET 반응식을 선택해 **`[SMILES]` ➔ `[Copy KET]`** 를 누른 뒤, PowerPoint로 이동해 **`[Paste Structure]`** ➔ **`[Batch Style]` ➔ `[Dark Mode]`** 를 실행하세요. 화살표와 촉매 주석까지 투명 흰색 선화로 즉시 전환됩니다.
:::

## 관련 가이드

- [**2. 화합물 표(SAR) & 실험 물성(SI)**](/ko/word/tables-properties) — 열 너비 고정 SAR 표 구성 및 `Molecule Info` 12종 물성 추출
- [**스타일 & 출력 상세 매뉴얼**](/ko/editor/publishing) — `Batch Style` 4종 프리셋 및 `Molecule Info` 5가지 복사 서식 전체 예시
- [**문서 공유 가이드**](/ko/guide/troubleshooting/document-sharing) — macOS·추가 기능 없는 PC에서의 `.docx` / `.pptx` / `.xlsx` 동작 방식
