# Word: 학술 논문 본문 & 요약 표(Table) 작성 가이드

학술 논문, 학위 논문, 연구 보고서 본문과 표(Table) 안에 고해상도 2D 화학 구조식·반응식 및 3D 단백질을 삽입하고, 클릭 한 번으로 논문 전체를 **학술지 흑백(ACS 스타일)**으로 통일하는 실무 가이드입니다.

<p align="center">
  <img src="/word-showcase.png" alt="func_SMILES2IMG Word Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🎯 워드에서 시작하는 3단계 방법 (수식 입력 없음: ① ➔ ② ➔ ③)

엑셀과 달리 워드에서는 셀 수식을 치지 않고 **상단 `[SMILES]` 리본 메뉴** 또는 단축키 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>로 바로 시작합니다:

<p align="center">
  <img src="/word-ribbon-guide.png" alt="Microsoft Word 상단 SMILES 리본 메뉴" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **① 상단 `[SMILES]` 탭 ➔ `[Insert / Edit Structure]` 클릭** (또는 키보드 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
2. **② 2D 구조·반응식 그리기 (또는 클립보드 붙여넣기):** 내장된 Ketcher 편집기에서 분자나 다단계 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 커서 위치에 인라인 그림으로 바로 들어갑니다.
   * *복사해 둔 SMILES가 있다면?* **`[Paste Structure]`** 버튼을 누르면 0.1초 만에 구조식 그림으로 변환됩니다!
3. **③ 언제든 더블클릭으로 제자리 수정:** 본문이나 **워드 표(Table) 안의 분자 그림을 더블클릭**하면 표의 열 너비나 레이아웃이 전혀 깨지지 않고 그 자리에서 바로 수정됩니다!

---

## 2. 투고 학술지 규정에 맞춰 문서 전체 스타일 통일하기 (`Batch Style`)

논문을 투고할 때 수십 개의 분자 그림을 하나하나 다시 그릴 필요 없이, 상단 **`[SMILES]` ➔ `[Batch Style]`** 버튼 하나로 문서 전체(또는 선택한 영역)의 구조식 스타일을 통일할 수 있습니다.

| 스타일 프리셋 | 변환 결과 | 추천 활용처 |
| :--- | :--- | :--- |
| **Standard Color** | 원소별 표준 컬러 + **투명 배경** | 컬러 학위 논문, 전자 문서(PDF), 연구 보고서 |
| **Publication B&W** | 고대비 **흑백 선화 (ACS 스타일)** | ACS, Wiley, Elsevier 등 인쇄 학술지 및 특허 명세서 투고 |
| **White Background** | 원소별 컬러 + **흰색 사각형 배경** | 배경 음영이 칠해진 워드 표(Table)나 강조 박스 내부 |
| **Dark Mode** | **흰색 결합선** + 투명 배경 | 어두운 배경의 표지 및 다크 테마 문서 |

---

## 3. 워드 표(Table)에 분자 물성 채우기 & Illustrator 벡터 복사

워드 문서에 삽입된 화학 구조식을 **우클릭 ➔ `Chemical Structure`** (또는 상단 리본 메뉴)로 선택하면 논문 작성 시간을 대폭 줄여주는 도구들을 쓸 수 있습니다.

<p align="center">
  <img src="/word-context-menu-guide.png" alt="Chemical Structure Context Menu in Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### 📋 실험 섹션(Experimental) 및 워드 표에 물성 바로 붙여넣기
* **개별 물성 원클릭 복사:** 물질명(Substance Name), 화학식(`C9H8O4`), 분자량(`180.16 g/mol`), 정밀 질량, InChIKey를 클릭 한 번으로 복사해 논문 본문에 붙여넣습니다.
* **워드 표 행(TSV)으로 한 번에 붙여넣기:** **`Copy Molecule Info` ➔ `Copy as Excel Row (TSV Columns)`** 메뉴를 클릭한 뒤 워드 표(Table)에 붙여넣으면 각 셀 칸에 맞춰 물성 값들이 자동으로 나뉘어 들어갑니다!
* **상세 정보 창 (`Detailed Molecule Info Window...`):** 분자량, 정밀 질량, LogP, TPSA, Lipinski 약물성 통과 여부를 한눈에 확인하고 마크다운 표나 JSON으로도 복사할 수 있습니다.

### 🖼️ Adobe Illustrator / ChemDraw용 무손실 벡터 복사 (`Copy SVG`)
* 상단 **`[Copy SVG]`** (또는 우클릭 **`Copy Vector SVG`**) 메뉴를 클릭하면 화질 저하가 전혀 없는 표준 벡터 그래픽(`SVG`)으로 복사되어 Illustrator, Inkscape 등에 바로 붙여넣고 편집할 수 있습니다.

---

## 4. 논문에 3D 단백질 구조 넣기 (`Mol*`)

1. 상단 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 버튼을 클릭합니다.
2. 4자리 PDB 코드(예: `1CRN`, `4HHB`, `6LU7`)를 입력해 다운로드하거나 내 PC의 구조 파일(`.pdb`, `.cif`, `.sdf`)을 엽니다.
3. 원하는 스타일(Cartoon 리본, 표면 등)과 각도를 맞춘 뒤:
   * **[3D Model (.glb)]:** 워드 문서 안에서 마우스로 360° 직접 돌려볼 수 있는 인터랙티브 3D 모델로 삽입합니다.
   * **[2D Picture (.png)]:** 논문 인쇄 및 PDF 제출용 고해상도 그림으로 삽입합니다.
4. 삽입된 3D 모델이나 그림 역시 **더블클릭**하면 작업하던 각도 그대로 다시 열려 언제든 수정할 수 있습니다.

---

## 5. 상단 `[SMILES]` 리본 메뉴 한눈에 보기

<p align="center">
  <img src="/word-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ `Insert / Edit Structure`:** 2D 분자 및 합성 반응식 스케처 실행 (선택된 분자가 있으면 즉시 수정 모드로 열림)
* **🧬 `3D Structure (Mol*)`:** 3D 단백질/고분자 뷰어 실행
* **📋 `Paste Structure`:** 드래그한 SMILES 텍스트나 클립보드 데이터를 0.1초 만에 구조식 그림으로 삽입
* **🖼️ `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** 선택한 구조식에서 화학 식별자 또는 무손실 SVG 벡터 복사
* **🔬 `Molecule Info`:** 분자량·화학식·약물성 확인 및 다양한 표 서식으로 복사
* **🎨 `Batch Style`:** 문서 내 모든 구조식을 컬러 / 논문용 흑백(ACS) / 흰색 배경 / 다크 모드로 일괄 변환

---

## 🚀 다음 단계로 이어서 보기

<div class="s2i-card-grid">
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 출판 스타일 · SVG · 물성 복사</div>
    <p class="s2i-card-desc">학술지 투고용 ACS 흑백 일괄 변환 프리셋 비교 및 5가지 물성 내보내기 서식</p>
  </a>
  <a href="../guide/powerpoint" class="s2i-card">
    <div class="s2i-card-title">📽️ PowerPoint 발표 슬라이드</div>
    <p class="s2i-card-desc">다크 모드 슬라이드용 투명 구조식과 발표 중 360° 회전하는 3D 단백질 모델</p>
  </a>
</div>

