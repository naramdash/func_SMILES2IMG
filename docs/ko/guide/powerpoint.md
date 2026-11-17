# PowerPoint 슬라이드 갤러리: 다크 모드 & 360° 3D 발표 슬라이드

> **PowerPoint 가이드 구성 방식 — 비주얼 슬라이드 갤러리:** 완성된 발표 슬라이드 패턴의 실제 구동 영상과 캡처 화면을 먼저 눈으로 확인하고, 마음에 드는 슬라이드를 3단계로 바로 따라 만들어 보세요.

<FeatureMedia id="powerpoint-overview" />

---

## 슬라이드 패턴 1: 다크 테마 4종 후보 물질 비교 카드 그리드

어두운 학회 발표 템플릿 위에서도 하얀 사각형 배경 없이 투명한 고대비 구조식 카드를 배치합니다.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. 슬라이드 배경을 짙은 네이비/슬레이트(`#0F172A`)로 설정하고 카드 박스를 배치합니다.
2. **`[SMILES] ➔ [Paste Structure]`** 또는 **`[Insert / Edit Structure]`**로 각 후보 물질을 삽입합니다.
3. **`[SMILES] ➔ [Batch Style] ➔ [Dark Mode]`**를 누르면 모든 구조식이 투명 배경의 흰색 결합선으로 통일됩니다.

---

## 슬라이드 패턴 2: 다크 모드 흰색 결합선 (투명 캔버스)

빔프로젝터 발표 시 어두운 배경 위에서도 원자와 결합선이 또렷하게 빛나는 다크 모드 스타일입니다.

<FeatureMedia id="powerpoint-slide-dark-mode" />

* **Ketcher 내부 설정:** 상단 바에서 **`[Background: Transparent]`**를 선택한 뒤 슬라이드에 적용하세요.
* **일괄 전환:** 슬라이드 위 도형을 선택하거나(선택 없으면 **발표 자료 전체**) **`[Batch Style] ➔ [Dark Mode]`**를 클릭합니다.

<BatchStylePreview />

---

## 슬라이드 패턴 3: 다단계 유기 합성 반응 메커니즘 슬라이드

반응 화살표(`➔`), `+`, 촉매 조건을 함께 배치하고 `Shape.Tags["KET"]`에 원본 데이터를 무손실 저장합니다.

<FeatureMedia id="powerpoint-slide-reaction" />

1. **`[SMILES] ➔ [Insert / Edit Structure]`** (<kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>)를 클릭합니다.
2. Ketcher에서 반응물·화살표·시약 텍스트를 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누릅니다.
3. 발표 직전 화살표나 치환기를 고쳐야 할 때는 **슬라이드 도형을 더블클릭**하면 카드 위치와 크기를 유지한 채 제자리 수정됩니다.

---

## 슬라이드 패턴 4: 발표 중 마우스로 360° 돌리는 3D 단백질 슬라이드 (`Mol*`)

여러 각도의 캡처 그림을 붙이는 대신, 슬라이드 위에서 직접 회전·확대되는 3D 단백질 모델을 시연합니다.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

1. 상단 리본 **`[SMILES] ➔ [3D Structure (Mol*)]`**를 클릭합니다.
2. 4자리 PDB 코드(`1CRN`, `4HHB`, `6LU7`, `1BNA`)를 입력해 **`Fetch`**하거나 로컬 `.pdb`/`.cif`/`.sdf` 파일을 엽니다.
3. **`3D Model (GLB)`** (발표 중 360° 회전 가능) 또는 **`2D Picture`** (`BG: Transparent`) 모드를 선택해 슬라이드에 삽입합니다.

---

## 슬라이드 패턴 5: 텍스트 상자 SMILES 0.1초 만에 구조식으로 바꾸기

슬라이드에 적어둔 SMILES 문자열이나 시약 이름을 에디터 창 없이 클릭 한 번으로 분자 도형으로 바꿉니다.

<FeatureMedia id="powerpoint-slide-textbox-quick-insert" />

1. SMILES 문자열(예: `CC(=O)Oc1ccccc1C(=O)O`)이나 시약명(`Aspirin`)이 적힌 텍스트 상자를 선택합니다.
2. 리본 **`[SMILES] ➔ [Paste Structure]`**를 클릭합니다.
3. 가로세로 비율이 잠긴(`msoTrue`) 고화질 분자 그림이 슬라이드에 즉시 배치됩니다.

---

## 슬라이드 패턴 6: 발표 Q&A 중 분자량·약물성 즉시 확인 (`Molecule Info`)

질의응답 중 청중이나 심사위원이 분자량·정밀질량·Lipinski 통과 여부를 물어볼 때 슬라이드 도형에서 바로 조회합니다.

<FeatureMedia id="powerpoint-slide-live-qa-info" />

1. 슬라이드의 분자 도형을 선택하고 **`[SMILES] ➔ [Molecule Info]`** (또는 우클릭 ➔ **`Chemical Structure ➔ Copy Molecule Info`**)를 클릭합니다.
2. **`Copy as Multi-Line List`**를 눌러 슬라이드 글머리 기호(불릿) 목록으로 바로 붙여넣거나, **`Detailed Molecule Info Window...`**로 17개 물성을 한눈에 확인합니다.

---

## 슬라이드 패턴 7: 발표 자료 전체 테마 원클릭 전환 (`Batch Style`)

다크 테마 발표용 슬라이드를 인쇄용 흰색 유인물로 바꾸거나, 반대로 라이트 슬라이드를 다크 모드로 일괄 변환합니다.

<FeatureMedia id="powerpoint-slide-batch-style" />

1. 특정 도형 몇 개만 선택하거나, 아무것도 선택하지 않은 상태에서(프레젠테이션 전체) **`[SMILES] ➔ [Batch Style]`**를 엽니다.
2. **`Standard Color`**, **`Publication B&W`**, **`White Background`**, **`Dark Mode`** 중 원하는 프리셋을 클릭합니다.

---

## PowerPoint 3단 우클릭 메뉴 & 학회 포스터 SVG 복사

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Microsoft PowerPoint 실제 3단 Chemical Structure 우클릭 메뉴" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* 슬라이드 분자 그림 우클릭 ➔ **`Chemical Structure ➔ Copy Vector SVG`**를 클릭하면 A0 대형 학회 포스터에서도 전혀 깨지지 않는 무손실 벡터가 복사됩니다.
* 리본 끝 **`[About func_SMILES2IMG]`** 버튼에서 버전 정보와 최신 업데이트를 확인할 수 있습니다.

<div class="s2i-card-grid">
  <a href="../guide/molstar" class="s2i-card">
    <div class="s2i-card-title">🧬 Mol* 3D 단백질 뷰어 상세 가이드</div>
    <p class="s2i-card-desc">추천 PDB 코드 4선과 더블클릭 카메라 각도 복원 방법 확인하기</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">📐 출판 스타일 · 벡터 SVG · 물성 복사</div>
    <p class="s2i-card-desc">학회 포스터용 SVG 복사 및 5가지 물성 내보내기 서식 자세히 보기</p>
  </a>
</div>
