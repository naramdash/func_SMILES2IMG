# PowerPoint: 다크 모드 발표 슬라이드 & 360° 3D 단백질

어두운 학회 발표 슬라이드 위에서도 하얀 사각형 여백 없이 투명 배경(흰색 선)으로 구조식을 배치하고, 발표 중 마우스로 360° 돌려보는 3D 단백질 모델을 삽입하는 실무 가이드입니다.

<p align="center">
  <img src="/powerpoint-showcase.png" alt="func_SMILES2IMG PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🎯 파워포인트에서 시작하는 3단계 방법 (수식 입력 없음: ① ➔ ② ➔ ③)

파워포인트에서는 셀 수식을 치는 대신 **상단 `[SMILES]` 리본 메뉴** 또는 단축키 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르면 즉시 시작됩니다:

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="Microsoft PowerPoint 상단 SMILES 리본 메뉴" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **① 상단 `[SMILES]` 탭 ➔ `[Insert / Edit Structure]` 클릭** (또는 키보드 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
2. **② 스케치 후 슬라이드 배치 (`Ctrl + Enter`):** Ketcher 편집기에서 분자나 합성 경로를 그리고(어두운 배경 슬라이드라면 상단 **`[배경: 투명]`** 선택!) <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 슬라이드 중앙에 즉시 배치됩니다.
   * *복사한 SMILES가 있다면?* **`[Paste Structure]`** 버튼 한 번으로 0.1초 만에 도형으로 들어갑니다!
3. **③ 더블클릭으로 위치·크기 유지한 채 수정:** 슬라이드 위의 분자 그림을 **더블클릭**하면 맞춰 둔 슬라이드 위치와 크기가 흔들리지 않고 제자리에서 바로 업데이트됩니다.

---

## 2. 다크 모드(흰색 선) vs 일반 배경 원클릭 비교 및 일괄 변환 (`Batch Style`)

어두운 발표 슬라이드에 하얀 네모 배경 캡처를 붙여넣지 마세요! 상단 **`[SMILES] ➔ [Batch Style]`**을 클릭하면 슬라이드 전체의 구조식을 **다크 모드(투명 배경 + 흰색 결합선)**, **표준 컬러**, **인쇄 유인물용 흑백(ACS)**으로 클릭 한 번에 통일합니다. 아래 버튼을 눌러 Before / After를 직접 비교해 보세요:

<BatchStylePreview />

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

