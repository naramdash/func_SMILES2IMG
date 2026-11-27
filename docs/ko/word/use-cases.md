# Word 활용 사례

연구실과 기업 R&D 실무에서 **`func_SMILES2IMG`** 를 Microsoft Word에 활용하는 5가지 대표 시나리오입니다. 학술지 투고 마감부터 특허 SAR 표, 실험 보고서(SI), 3D 단백질 보고서, 장문 학위논문까지 바로 적용해 보세요.

| 시나리오 | 해결하는 작업 | 핵심 기능 |
| :--- | :--- | :--- |
| [**1. 학술지 투고 · 흑백 통일**](#journal-manuscript) | 컬러 초안 전체를 순수 흑백 선화로 일괄 변환하고 TOC 벡터 SVG 추출 | `[Batch Style]`, `[Copy SVG]` |
| [**2. 특허 명세서 · SAR 화합물 표**](#patent-sar) | 열 너비 고정 상태에서 구조식 변환 및 계산된 물성 표 채움 | `[Paste Structure]`, `[Molecule Info]` |
| [**3. 합성 실험 보고서 · SI 작성**](#lab-report-si) | 무손실 편집 가능한 `KET` 반응식 저장 및 SI 특성 요약 목록 복사 | `[Insert / Edit (KET)]`, `Multi-Line List` |
| [**4. 3D 단백질 · 결합 포켓 보고서**](#structural-reports) | PDB·로컬 `.pdb` 결합 포켓(`PNG` / 3D `.glb`) 삽입 및 카메라 각도 재편집 | `[3D Structure (Mol*)]` |
| [**5. 장문 학위논문 · 인쇄본 분기**](#thesis-dissertation) | Word 캡션·상호 참조를 유지한 채 심사본(컬러)과 제본본(흑백) 생성 | 더블클릭 편집 + `[Batch Style]` |

---

## 1. 학술지 투고 · 흑백 통일 {#journal-manuscript}

내부 검토용 초안은 보기 편한 IUPAC 컬러로 작성해 Mac 공동저자와 주고받은 뒤, 투고 전 원고 전체를 ACS 흑백 선화로 바꾸고 Illustrator용 벡터 TOC Graphic을 추출할 수 있습니다.

<FeatureMedia id="word-ch6-publication-finish" />

1. **초안은 컬러로 작성:** 본문과 표에 기본 IUPAC 컬러로 인라인 구조식과 `KET` 반응식을 넣습니다.
2. **투고 전 흑백 일괄 통일:** 그림 선택 없이 본문에 커서를 둔 채 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** 를 클릭하면 모든 2D 구조식·반응식이 열 너비 변경 없이 순수 흑백으로 바뀝니다(`Mol*` 3D 그림은 컬러 유지).
3. **Illustrator용 TOC 벡터 추출:** 대표 반응식을 선택하고 **`[SMILES]` ➔ `[Copy SVG]`** ➔ Illustrator/Inkscape에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣어 선 굵기까지 자유롭게 편집합니다.
4. **Mac 공동저자 · 심사 포털 공유:** 표준 고해상도 PNG로 저장되므로 애드인 없는 Mac·웹 Word에서도 그대로 보이며, 대체 텍스트(Alt Text) 태그 덕분에 다음 교정 시 Windows에서 다시 더블클릭 편집이 가능합니다.

* **관련 가이드:** [**4. 투고 마감 · 벡터 SVG · 공동저자 공유**](/ko/word/publishing) · Playbook **6쪽**

---

## 2. 특허 명세서 · SAR 화합물 표 {#patent-sar}

특허 청구항이나 논문 SAR 요약 표에서 치환기를 수정할 때마다 옆 열이 찌그러지거나 상단 고리가 잘리는 문제, 화학식·분자량·정밀질량(`Da`) 수기 오타를 한 번에 해결합니다.

<FeatureMedia id="word-ch3-compound-table" />

1. **표 레이아웃 고정:** **표 우클릭 ➔ 자동 맞춤 ➔ 열 너비 고정**을 켜고, 구조식 셀의 단락 줄 간격이 `고정값`이 아닌 **`1줄`** 또는 **`배수`** 인지 확인합니다.
2. **텍스트 목록 제자리 변환:** 표 셀에 `Aspirin`, `Caffeine`, `"50-78-2"` 또는 SMILES를 붙여넣고 텍스트를 선택해 **`[SMILES]` ➔ `[Paste Structure]`** 를 누릅니다.
3. **계산된 물성 자동 채움:** 구조식(또는 텍스트)을 선택해 **`[Molecule Info]` ➔ `Copy as Excel Row (TSV Columns)`** 로 행 데이터를 붙이거나 화학식·정밀질량을 개별 복사합니다.
4. **치환기 제자리 수정:** 표 셀 안의 구조식을 **더블클릭**해 수정해도 잠긴 열 너비 안에서 비율을 유지하며 즉시 교체됩니다.

* **관련 가이드:** [**2. 화합물 표(SAR) & 실험 물성(SI)**](/ko/word/tables-properties) · Playbook **3쪽**

---

## 3. 합성 실험 보고서 · SI 작성 {#lab-report-si}

다단계 합성 경로의 화살표·촉매·온도·수율 배치까지 Word 문서 안에 무손실 데이터로 보관하고, Supporting Information(SI)용 화합물 물성·식별자 블록을 오타 없이 복사해 넣습니다.

<FeatureMedia id="word-ch4-reaction-scheme" />

1. **무손실 `KET` 반응식 삽입:** **`[SMILES]` ➔ `[Insert / Edit Structure]`** 에서 반응물, 화살표(`➔`), `+`, 조건 텍스트(`T` 도구: `H₂SO₄, 85 °C, 92%`)를 그려 넣습니다. 대체 텍스트(`KET:...`)에 전체 좌표가 보존되어 수주 뒤 수율 수정 시에도 더블클릭으로 원래 배치가 복원됩니다.
2. **SI 특성 블록 생성:** 화합물 그림이나 본문의 SMILES 텍스트를 선택하고 **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** 를 눌러 화학식, MW, 정밀질량(`Da`), LogP, TPSA, SMILES, InChI, InChIKey를 붙여넣습니다.
3. **투고 시스템 식별자 추출:** 학술지 포털 제출용 문자열은 **`[Copy SMILES]`** · **`[Copy InChI]`** 로 즉시 복사합니다.

* **관련 가이드:** [**1. 2D 구조식 · 반응식 작성**](/ko/word/structures) · Playbook **2쪽 & 4쪽**

---

## 4. 3D 단백질 · 결합 포켓 보고서 {#structural-reports}

연구비 제안서나 프로젝트 보고서에 PDB 공개 구조 또는 사내 도킹 파일(`.pdb`, `.cif`) 결합 포켓 그림을 넣고, 리뷰 의견에 맞춰 카메라 각도를 그 자리에서 바로 다듬어 교체합니다.

<FeatureMedia id="word-ch5-edit-and-molstar" />

1. **본문 PDB 코드 바로 열기:** 원고의 `1CRN`·`6LU7` 글자를 드래그 선택하고 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 를 클릭하거나 **`Open File...`** 로 사내 `.pdb`/`.cif`/`.sdf` 파일을 오프라인으로 엽니다.
2. **목적별 포맷 선택:** 투고·PDF 제출본은 **`2D Picture`** (`BG: White` / `Transparent`), 사내 전자 보고서(Word 2019+/365)는 문서 안에서 360° 돌려볼 수 있는 **`3D Model (GLB)`** 를 선택합니다.
3. **카메라 각도 유지 재편집:** 삽입된 그림을 **더블클릭**하면 `CustomXMLParts`에 저장된 카메라 각도와 배색 그대로 Mol\*이 다시 열립니다.

::: tip 3종 앱 연계 워크플로우
[**Excel 활용 사례**](/ko/excel/use-cases)에서 후보를 선별하고 ➔ Word 논문 표·반응식을 작성한 뒤 ➔ **`[Copy KET]`** / **`[Copy SMILES]`** 로 [**PowerPoint 활용 사례**](/ko/powerpoint/use-cases)의 다크모드 슬라이드로 바로 옮길 수 있습니다.
:::

* **관련 가이드:** [**3. 3D 단백질 Figure (`Mol*`)**](/ko/word/3d-figures) · Playbook **5쪽**

---

## 5. 장문 학위논문 · 인쇄본 분기 {#thesis-dissertation}

100페이지가 넘는 석·박사 학위논문에서도 Word **캡션(`Scheme 3.2`)** · 본문 **상호 참조** · **그림/반응식 목차**를 깨뜨리지 않고 수정할 수 있으며, 심사본(컬러 PDF)과 도서관 제본본(흑백 인쇄) 두 버전을 손쉽게 분기합니다.

<FeatureMedia id="word-usecase-thesis-captions" />

1. **캡션 · 상호 참조 무결성 유지:** 인라인 구조식 아래 Word 기본 **캡션 삽입**을 붙여 두세요. **더블클릭** 재편집이나 **`[Batch Style]`** 은 문단 내 그림 위치만 제자리 교체하므로 자동 번호와 목차가 유지됩니다.
2. **컬러 심사본 ↔ 흑백 제본본 일괄 전환:** 컬러 마스터 `.docx`로 심사본 PDF를 출력하고, 제본용 사본에서 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** 를 실행해 전체 흑백본을 만듭니다.
3. **수개월 뒤에도 원본 파일 찾기 불필요:** `SMILES` · `KET` 원본 좌표가 `.docx` 내부에 저장되어 있어 예전 PC의 `.cdx` 파일 없이도 더블클릭만으로 바로 수정합니다.

* **관련 가이드:** [**1. 2D 구조식 · 반응식 작성**](/ko/word/structures) · [**4. 투고 마감 · 벡터 SVG · 공동저자 공유**](/ko/word/publishing)

---

## 다음 단계

- [**6페이지 Word Playbook (.docx) 실습하기**](/ko/word/playbook) — `Paste Structure`, 고정 SAR 표, KET 반응식, `1CRN` 3D 그림, ACS 흑백 전환을 Word에서 직접 체험해 보세요.
- [**설치 가이드**](/ko/guide/install) — Word·Excel·PowerPoint 통합 설치 명령어를 확인하세요.
