# Word 6장 논문 튜토리얼: 원고 한 편 처음부터 투고까지 완성하기

> **Word 가이드 구성 방식 — 연속 스토리 튜토리얼:** 하나의 연구 논문 원고(*Synthesis and Structural Characterization of Aspirin Analogs*)를 1장 서론 삽입부터 6장 학술지 투고 마감까지 순서대로 완성해 나갑니다. 특정 기능만 빠르게 찾으려면 아래 **빠른 작업 색인(Task Index)**을 누르세요.

<FeatureMedia id="word-overview" />

## 빠른 작업 색인 (하고 싶은 일로 바로 이동)

| Word에서 하고 싶은 작업 | 이동할 장(Chapter) | 리본 버튼 / 단축키 |
| :--- | :---: | :--- |
| 본문에 새 2D 화학 구조식 직접 그려 넣기 | [1장](#제1장-서론-본문에-2d-구조식-삽입하기) | <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd> 또는 `[Insert / Edit Structure]` |
| 본문에 적은 SMILES 글자를 0.1초 만에 그림으로 바꾸기 | [2장](#제2장-본문-smiles-텍스트를-0-1초-만에-그림으로-변환) | 텍스트 선택 ➔ `[Paste Structure]` |
| 화합물 요약 표(Table)에 구조식과 분자량·화학식 채우기 | [3장](#제3장-화합물-요약-표-작성-molecule-info-복사) | 우클릭 ➔ `Copy Molecule Info` |
| 다단계 합성 반응식(Scheme) 넣고 화살표 보존하기 | [4장](#제4장-다단계-합성-반응식-삽입-무손실-ket-보존) | Ketcher 반응 화살표(`➔`) 작도 ➔ 적용 |
| 표 안 그림 제자리 수정 & 3D 단백질 Figure 삽입 | [5장](#제5장-그림-더블클릭-수정-3d-단백질-figure-삽입-mol) | 그림 **더블클릭** 또는 `[3D Structure (Mol*)]` |
| 투고 전 원고 전체 ACS 흑백 통일 & Illustrator SVG 복사 | [6장](#제6장-투고-전-acs-흑백-일괄-통일-svg-내보내기) | `[Batch Style]` 및 `[Copy SVG]` |

---

## 제1장: 서론 본문에 2D 구조식 삽입하기

1장에서는 논문 서론 문단 안에 아스피린(Aspirin) 구조식을 글자 흐름에 맞춰 깔끔한 인라인 이미지로 넣습니다.

<FeatureMedia id="word-ch1-insert-ketcher" />

1. 구조식이 들어갈 문장 끝에 커서를 둡니다.
2. 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`**를 클릭하거나 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누릅니다.
3. 오프라인 **Ketcher 2D 편집기**에서 분자를 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 커서 위치에 즉시 삽입됩니다.

---

## 제2장: 본문 SMILES 텍스트를 0.1초 만에 그림으로 변환

실험노트나 문서에 적혀 있는 SMILES 문자열을 편집기 창을 열지도 않고 마우스로 긁어서 곧바로 구조식 그림으로 치환합니다.

<FeatureMedia id="word-ch2-quick-insert" />

1. 문서 본문에 있는 SMILES 문자열(예: `Cn1cnc2c1c(=O)n(C)c(=O)n2C`)을 드래그해 선택합니다.
2. 리본 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure ➔ Paste Structure`**)를 클릭합니다.
3. 선택한 텍스트가 비율이 고정된 고화질 인라인 구조식으로 즉시 바뀝니다. *(텍스트를 선택하지 않고 누르면 클립보드의 SMILES/KET를 삽입합니다.)*

---

## 제3장: 화합물 요약 표 작성 & `Molecule Info` 물성 복사

3장에서는 워드 표(Table) 셀 안에 후보 화합물을 배치하고, 화학식·분자량·LogP를 클릭 한 번으로 채웁니다.

<FeatureMedia id="word-ch3-compound-table" />

1. 워드 표의 구조식 셀에 커서를 두고 **`[Insert / Edit Structure]`** 또는 **`[Paste Structure]`**로 구조식을 삽입합니다.
2. 삽입된 구조식을 선택하고 **`[SMILES] ➔ [Molecule Info]`** (또는 우클릭 ➔ **`Chemical Structure ➔ Copy Molecule Info`**)를 클릭합니다.
3. **`Copy as Excel Row (TSV Columns)`**를 선택하면 17개 물성이 탭으로 나뉘어 워드 표의 여러 셀에 한 번에 들어가며, **`Detailed Molecule Info Window...`**에서는 개별 값, 마크다운 표, JSON을 선택 복사할 수 있습니다.

---

## 제4장: 다단계 합성 반응식 삽입 (무손실 KET 보존)

4장에서는 살리실산 + 아세트산 무수물 ➔ 아스피린 합성 경로(Scheme 1)를 삽입합니다.

<FeatureMedia id="word-ch4-reaction-scheme" />

* Ketcher에서 반응 화살표(`➔`), `+` 기호, 촉매·온도 주석을 함께 그려 적용하면 그림 속성(`AlternativeText`)에 전체 KET 좌표가 무손실 보존됩니다.
* `.docx` 파일을 저장하고 공동 저자와 주고받은 뒤에도, **그림을 더블클릭**하면 화살표 각도와 주석 위치 그대로 다시 열려 수정할 수 있습니다.

---

## 제5장: 그림 더블클릭 수정 & 3D 단백질 Figure 삽입 (`Mol*`)

표 안에 넣은 구조식의 치환기를 고치거나, 약물이 결합하는 3D 단백질 구조를 논문의 Figure로 추가합니다.

<FeatureMedia id="word-ch5-edit-and-molstar" />

1. 본문이나 표 셀 안의 **2D 그림을 더블클릭**하면 열 너비나 문단 배치가 망가지지 않고 제자리에서 크기 비율(`LockAspectRatio = msoTrue`)을 유지한 채 수정됩니다.
2. 3D 단백질 Figure를 넣으려면 **`[SMILES] ➔ [3D Structure (Mol*)]`**를 클릭하고 `1CRN`, `6LU7` 등 4자리 PDB 코드를 불러온 뒤, **3D 모델(`.glb`)** 또는 **고해상도 2D 그림(`.png`)**으로 삽입합니다.

---

## 제6장: 투고 전 ACS 흑백 일괄 통일 & SVG 내보내기

학술지(ACS, Wiley, Elsevier, RSC) 투고 직전, 원고 전체에 흩어진 모든 분자 그림을 한 번에 논문 규격으로 통일합니다.

<FeatureMedia id="word-ch6-publication-finish" />

1. **`[SMILES] ➔ [Batch Style] ➔ [Publication B&W]`**를 클릭하면 선택한 영역(선택이 없으면 **원고 전체**)이 순수 흑백 선화로 통일됩니다. *(음영 처리된 표 셀에는 **`White Background`** 프리셋을 추천합니다.)*
2. Adobe Illustrator나 Inkscape에서 세밀하게 다듬을 Scheme은 선택 후 **`[Copy SVG]`**로 무손실 벡터 경로를 복사해 붙여넣습니다.

---

## Word 상단 `[SMILES]` 리본 & 우클릭 메뉴 참조

<p align="center">
  <img src="/word-context-menu-guide.png" alt="Microsoft Word 실제 3단 Chemical Structure 우클릭 메뉴" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **`Insert / Edit Structure`** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>), **`3D Structure (Mol*)`**, **`Paste Structure`**
* **`Copy SMILES`**, **`Copy InChI`**, **`Copy KET`**, **`Copy SVG`**
* **`Molecule Info`**, **`Batch Style`**, **`About func_SMILES2IMG`**

<div class="s2i-card-grid">
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 출판 스타일 · SVG · 5가지 물성 내보내기</div>
    <p class="s2i-card-desc">Batch Style 4종 인터랙티브 비교 및 TSV·마크다운·JSON 물성 복사 가이드</p>
  </a>
  <a href="../guide/powerpoint" class="s2i-card">
    <div class="s2i-card-title">📙 다음 앱: PowerPoint 슬라이드 갤러리</div>
    <p class="s2i-card-desc">논문 속 분자와 3D 단백질을 다크 모드 발표 슬라이드로 옮겨 보세요</p>
  </a>
</div>
