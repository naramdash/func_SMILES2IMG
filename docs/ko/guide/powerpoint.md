# Microsoft PowerPoint 가이드

학회 발표, 연구실 세미나, 강의 슬라이드에 고해상도 2D 화학 구조식과 합성 반응식을 배치하고, 다크 모드 슬라이드 변환 및 360° 회전 가능한 3D 단백질 모델을 활용하는 방법을 안내합니다.

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 파워포인트 발표 자료 제작 시 3가지 핵심 장점
1. **다크 모드 & 투명 배경 완벽 지원:** 어두운 발표 템플릿 위에서도 보기 싫은 흰색 사각형 테두리 없이 깔끔한 투명 배경(또는 흰색 선)으로 들어갑니다.
2. **더블클릭 수정 시 슬라이드 위치·크기 그대로 유지:** 발표 직전 분자 구조나 반응 조건을 **더블클릭**해 고쳐도, 맞춰 놓은 슬라이드 좌표와 크기가 전혀 틀어지지 않습니다.
3. **발표 중 360° 돌려보는 3D 단백질:** 정적인 캡처 사진 대신 슬라이드 쇼 도중 마우스로 직접 회전·확대할 수 있는 실제 3D 단백질(`.glb`)을 넣을 수 있습니다.
:::

---

## 1. 슬라이드에 구조식·반응식 배치하기 (3가지 방법)

### 방법 ① 직접 그려서 슬라이드 중앙에 넣기 (`Insert / Edit Structure`)
1. 파워포인트 슬라이드 화면에서 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** 버튼을 클릭합니다 (또는 빈 슬라이드 우클릭 ➔ **`Chemical Structure` ➔ `Insert / Edit Structure (Ketcher)...`**).
2. **Ketcher 2D 화학 편집기**에서 분자 구조나 다단계 합성 경로(화살표, 시약, 촉매 조건)를 그립니다.
   * *팁: 어두운 배경 슬라이드라면 편집기 상단에서 **`[배경: 투명 (Background: Transparent)]`** 옵션을 선택하세요.*
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 현재 슬라이드 중앙에 선명한 도형으로 즉시 배치됩니다!

### 방법 ② 텍스트 상자 안의 SMILES나 클립보드 데이터를 즉시 도형으로 넣기 (`Paste Structure`)
* **텍스트 상자 변환:** 슬라이드 텍스트 상자에 적힌 SMILES 문자열을 선택하고 상단 **`[SMILES]` ➔ `[Paste Structure]`** 버튼을 누르면 0.1초 만에 분자 도형으로 바뀝니다.
* **클립보드에서 바로 붙여넣기:** 복사해 둔 SMILES가 있다면 **`[SMILES]` ➔ `[Paste Structure]`** (또는 빈 슬라이드에서 **우클릭 ➔ `Chemical Structure` ➔ `Paste Structure from Clipboard`**) 메뉴를 눌러 바로 넣을 수 있습니다.

---

## 2. 다크 모드 & 발표 템플릿에 맞춰 전체 슬라이드 통일하기 (`Batch Style`)

발표장 환경(밝은 배경 vs 어두운 OLED 화면 vs 인쇄용 유인물)에 맞춰 슬라이드 전체의 화학 구조식을 클릭 한 번으로 바꿀 수 있습니다. 상단 **`[SMILES]` ➔ `[Batch Style]`** 버튼을 클릭하세요:

| 스타일 프리셋 | 슬라이드 변환 결과 | 추천 활용처 |
| :--- | :--- | :--- |
| **Standard Color** | 원소별 표준 컬러 + **투명 배경** | 밝은 배경 또는 그라데이션 발표 슬라이드 |
| **Dark Mode** | **선명한 흰색 결합선** + 투명 배경 | **어두운(Dark) 학회 발표 템플릿** 및 빔프로젝터 발표 |
| **Publication B&W** | 고대비 **흑백 선화 (ACS 스타일)** | 세미나 인쇄 유인물 및 흑백 자료집 |
| **White Background** | 원소별 컬러 + **불투명 흰색 배경** | 사진이나 복잡한 패턴이 깔린 슬라이드 배경 위 |

---

## 3. 발표 중 마우스로 360° 돌려보는 3D 단백질 넣기 (`Mol*`)

단백질-리간드 결합 부위나 3차 구조를 설명할 때 여러 각도의 캡처 사진을 여러 장 넣을 필요가 없습니다.

1. 상단 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 버튼을 클릭합니다.
2. 4자리 PDB 코드(예: `1CRN`, `4HHB`, `6LU7`)를 입력해 불러오거나 로컬 구조 파일을 엽니다.
3. 원하는 스타일(Cartoon, Surface, Ball & Stick)로 꾸민 뒤 상단에서 **`[3D Model (.glb)]`** 버튼을 눌러 삽입합니다.
4. 슬라이드 편집 화면이나 **슬라이드 쇼 발표 중에** 단백질 중앙의 3D 회전 핸들을 마우스로 드래그하면 청중 앞에서 360° 자유롭게 돌려가며 설명할 수 있습니다!
   *(고해상도 평면 그림이 필요할 때는 **`[2D Picture (.png)]`** 버튼으로 삽입할 수도 있습니다.)*

---

## 4. 슬라이드 글머리 기호(불릿)로 분자 정보 붙여넣기 & SVG 복사

슬라이드 위의 분자 도형을 **우클릭 ➔ `Chemical Structure`** 메뉴로 선택하면 발표 준비에 유용한 기능들을 바로 쓸 수 있습니다:

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **발표 슬라이드 불릿 목록으로 붙여넣기:** **`Copy Molecule Info` ➔ `Copy as Multi-Line List`** 항목을 클릭한 뒤 슬라이드 텍스트 상자에 붙여넣으면 물질명·화학식·분자량이 깔끔한 글머리 기호 목록으로 바로 들어갑니다!
* **크기를 아무리 키워도 안 깨지는 벡터 복사 (`Copy SVG`):** 포스터 발표(학회 포스터)나 Illustrator 편집을 위해 상단 **`[Copy SVG]`** (또는 우클릭 **`Copy Vector SVG`**) 항목을 누르면 무손실 벡터 그래픽으로 복사됩니다.

---

## 5. 상단 `[SMILES]` 리본 메뉴 한눈에 보기

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ `Insert / Edit Structure`:** 2D 분자 및 합성 반응식 스케처 실행 (선택된 분자 도형이 있으면 즉시 수정 모드로 열림)
* **🧬 `3D Structure (Mol*)`:** 인터랙티브 3D 단백질/고분자 뷰어 실행
* **📋 `Paste Structure`:** 선택한 텍스트 상자의 SMILES나 클립보드 데이터를 0.1초 만에 슬라이드 도형으로 배치
* **🖼️ `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** 화학 식별자 또는 학회 포스터·Illustrator용 무손실 SVG 벡터 복사
* **🔬 `Molecule Info`:** 슬라이드 불릿 목록용 물성 복사 및 상세 물성 창 열기
* **🎨 `Batch Style`:** 프레젠테이션 전체 구조식을 다크 모드 / 컬러 / 흑백 / 흰색 배경으로 원클릭 변환

---

## 🚀 다음 단계로 이어서 보기

<div class="s2i-card-grid">
  <a href="../guide/molstar" class="s2i-card">
    <div class="s2i-card-title">🧬 Mol* 3D 단백질 · 결정 뷰어</div>
    <p class="s2i-card-desc">추천 PDB 코드 4선(1CRN, 4HHB, 6LU7, 1BNA)으로 3D 단백질 리본 및 결합 포켓 시각화</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 출판 스타일 · SVG · 물성 복사</div>
    <p class="s2i-card-desc">학회 포스터용 무손실 SVG 벡터 복사와 다크 모드 일괄 변환 시뮬레이터</p>
  </a>
</div>

