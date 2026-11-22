# PowerPoint 활용 사례

**`func_SMILES2IMG`** 로 Microsoft PowerPoint 학회 구두 발표, 주간 랩미팅 SAR 비교, 대형 A0 포스터, 합성 세미나 발표까지 완성하는 5가지 실전 시나리오입니다.

| 시나리오 | 해결하는 작업 | 핵심 기능 |
| :--- | :--- | :--- |
| [**1. 학회 발표 · 다크 모드 & 3D 모션**](#conference-talk) | 흰 사각형 없는 다크 모드 적용 및 `모핑(Morph)` 기반 3D 단백질 줌인 연출 | `[Batch Style]`, `[3D Structure (Mol*)]` |
| [**2. 주간 랩미팅 4-카드 SAR 비교**](#sar-review) | 텍스트 상자를 `0.1초` 만에 구조식으로 바꾸고 카드 정렬 유지한 채 치환기 수정 | `[Paste Structure]`, 더블클릭 |
| [**3. 대형 A0 포스터 · 흑백 유인물**](#a0-posters) | 인쇄 시 깨지지 않는 벡터 `SVG` 복사 및 발표 덱 사본의 1초 흑백 유인물 변환 | `[Copy SVG]`, `[Publication B&W]` |
| [**4. 합성 메커니즘 세미나 · Q&A 대비**](#qa-seminar) | Word·Excel `KET` 반응식 3초 재활용 및 발표자 노트 Q&A 물성 요약 세팅 | `[Copy KET]`, `Multi-Line List` |
| [**5. 단계별 반응식 SVG 애니메이션**](#svg-step-animation) | SVG를 개별 PowerPoint 도형으로 분해해 반응 부위 강조 및 클릭별 순차 전개 | `[Copy SVG]` ➔ `도형으로 변환` |

---

## 1. 학회 발표 · 다크 모드 & 3D 모션 {#conference-talk}

정적 PyMOL 캡처나 코덱 위험이 있는 동영상 대신, 어두운 템플릿 위에 투명 흰색 선화 2D 구조식과 **슬라이드 안에서 직접 회전·줌인되는 네이티브 3D `.glb` 단백질**을 배치해 어느 학회장 연단 PC·Mac에서도 안전하게 발표할 수 있습니다.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

1. **덱 전체 다크 모드 1초 통일:** 구조식을 넣은 뒤 빈 슬라이드 배경을 클릭하고 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`** 를 실행합니다.
2. **네이티브 3D `.glb` 단백질 삽입:** **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 에서 PDB ID(`1CRN`, `6LU7`) 또는 로컬 `.pdb`/`.cif` 파일을 불러와 **`3D Model (GLB)`** 로 삽입합니다.
3. **PowerPoint 자체 3D 애니메이션 적용:**
   - **한 슬라이드 내 3D 모션 (`애니메이션`):** `.glb` 모델을 선택해 **`회전판(Turntable)`** (360° 자전), **`스윙(Swing)`** (결합 포켓 중심 좌우 흔들기), **`도착(Arrive)`** 효과를 줍니다.
   - **슬라이드 전환 줌인 (`전환 ➔ 모핑`):** 슬라이드를 복제해 두 번째 장의 `.glb` 모델을 활성 부위 포켓에 바짝 줌인하고 **모핑(Morph)** 을 켜면 카메라가 부드럽게 날아 들어갑니다.
4. **발표장 PC · Mac 100% 호환:** 표준 내장 개체이므로 `func_SMILES2IMG` 미설치 PC와 macOS에서도 3D 회전과 모핑이 그대로 재생됩니다.

* **관련 가이드:** [**3. 360° 3D 단백질 발표 (`Mol*` & 모핑)**](/ko/powerpoint/3d-proteins) · [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) · Playbook **Slide 1 & 6**

---

## 2. 주간 랩미팅 4-카드 SAR 비교 {#sar-review}

매주 반복되는 유도체 4~8종 비교 슬라이드를 몇 분 만에 조립하고, 다음 주 회의 전에는 치환기 바뀐 카드만 더블클릭해 카드 정렬을 흐트러뜨리지 않고 교체합니다.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. **텍스트 상자 `0.1초` 변환:** 4개 카드에 SMILES나 화합물명(`Aspirin`, `Caffeine`)이 적힌 텍스트 상자를 배치하고 테두리를 선택해 **`[SMILES]` ➔ `[Paste Structure]`** 를 누릅니다.
2. **1클릭 스타일 통일:** **`[SMILES]` ➔ `[Batch Style]`** 로 모든 카드의 선 굵기와 투명 배경을 통일합니다 *(`Ctrl+G` 그룹 도형은 먼저 해제 후 실행)*.
3. **오타 없는 물성 불릿 붙이기:** 각 후보를 선택해 **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** 를 눌러 화학식·MW·LogP·TPSA·Lipinski 통과 여부를 카드 하단에 붙입니다.
4. **더블클릭 제자리 수정:** 수정할 구조식만 **더블클릭**해 Ketcher에서 고치면 카드 중앙 위치와 크기 그대로 교체됩니다.

* **관련 가이드:** [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) · [**1. 2D 구조식 · 반응식 슬라이드**](/ko/powerpoint/slides-structures) · Playbook **Slide 3 & 4**

---

## 3. 대형 A0 포스터 · 흑백 유인물 {#a0-posters}

가로·세로 1미터짜리 **A0/A1 학회 포스터(300 DPI)** 인쇄에서도 계단 현상 없이 칼날처럼 선명한 벡터 그래픽을 사용하고, 같은 발표 파일에서 인쇄용 흑백 PDF 유인물까지 1초 만에 뽑아냅니다.

<FeatureMedia id="powerpoint-slide-batch-style" />

1. **A0 캔버스에 무한 해상도 벡터 `SVG` 복사:** 2D 구조식이나 반응식(또는 텍스트 상자의 SMILES·화합물명 문자열)을 선택하고 **`[SMILES]` ➔ `[Copy SVG]`** 를 누른 뒤, 대형 PowerPoint 포스터(Office 2019+/365) 또는 **Adobe Illustrator / Inkscape**에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣습니다.
2. **1초 흑백 PDF 유인물 변환:** 다크 모드 발표 파일 사본을 열어 아무 도형도 선택하지 않은 채 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** 를 실행한 뒤 PDF로 내보냅니다.

* **관련 가이드:** [**4. A0 포스터(SVG) · Q&A · 발표장 공유**](/ko/powerpoint/publishing-qa) · [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) · Playbook **Slide 7 & 8**

---

## 4. 합성 메커니즘 세미나 · Q&A 대비 {#qa-seminar}

Word 논문이나 Excel 시트에 이미 만들어 둔 `KET` 반응식을 다시 그리지 않고 세미나 슬라이드로 가져오며, 청중 Q&A에서 정밀질량(`Da`)·`Fsp3`·`TPSA` 질문이 나와도 **발표자 노트**에서 바로 확인해 답합니다.

<FeatureMedia id="powerpoint-slide-reaction" />

1. **논문 반응식 3초 재활용:** Word 또는 Excel의 `KET` 반응식을 선택해 **`[SMILES]` ➔ `[Copy KET]`** 를 누른 뒤 PowerPoint에서 **`[Paste Structure]`** ➔ **`[Batch Style]` ➔ `[Dark Mode]`** 를 적용합니다 *(PowerPoint 애니메이션은 화학 편집이 끝난 뒤 마지막에 지정)*.
2. **발표자 노트 Q&A 세팅:** 발표 전 주요 물질을 선택해 **`[Molecule Info]` ➔ `[Copy as Multi-Line List]`** 를 눌러 **발표자 노트**에 붙여두면 발표자 보기에서 청중에게 보이지 않게 수치를 읽을 수 있습니다.
3. **17종 전체 물성 창 띄우기:** 세미나 Q&A 중에는 <kbd>Esc</kbd>로 쇼를 나와 구조식을 클릭하고 **`[Molecule Info]` ➔ `[Detailed Molecule Info Window...]`** 로 전체 물성을 함께 확인합니다.

::: tip 3종 앱 연계 워크플로우
[**Excel 활용 사례**](/ko/excel/use-cases)에서 후보를 선별하고 ➔ [**Word 활용 사례**](/ko/word/use-cases)에서 논문 표·반응식을 작성한 뒤 ➔ 완성된 `KET` · `SMILES` 를 그대로 PowerPoint 슬라이드와 A0 포스터로 옮겨 마무리하세요.
:::

* **관련 가이드:** [**1. 2D 구조식 · 반응식 슬라이드**](/ko/powerpoint/slides-structures) · [**4. A0 포스터(SVG) · Q&A · 발표장 공유**](/ko/powerpoint/publishing-qa) · Playbook **Slide 5 & 7**

---

## 5. 단계별 반응식 SVG 애니메이션 {#svg-step-animation}

강의나 메커니즘 세미나에서 다단계 합성 경로를 **반응물 ➔ 화살표·촉매 ➔ 생성물** 순으로 클릭마다 순차 등장시키거나, 특정 반응 부위 결합선만 강조 색으로 칠해 가독성을 높입니다 (Office 2019–365).

1. **먼저 화학 구조 완성:** **Ketcher 2D**에서 반응식과 촉매 주석을 최종 수정한 뒤 도형(또는 SMILES 텍스트)을 선택합니다.
2. **네이티브 SVG 붙여넣기:** **`[SMILES]` ➔ `[Copy SVG]`** 클릭 후 슬라이드에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣습니다.
3. **개별 PowerPoint 도형 분해:** 붙여넣은 SVG 그래픽을 **우클릭 ➔ 그룹 ➔ 그룹 해제** (또는 **그래픽 서식 ➔ 도형으로 변환**)하면 모든 결합선·화살표·원자 문자가 독립된 벡터 도형으로 분리됩니다.
4. **강조 채색 & 순차 애니메이션:** 반응 부위 결합선 색을 바꾸고 1·2·3단계 요소를 각각 다시 그룹(<kbd>Ctrl</kbd> + <kbd>G</kbd>)으로 묶은 뒤 **나타내기** / **닦아내기(왼쪽에서)** 애니메이션을 순서대로 지정하세요.

::: tip 백업용 숨김 슬라이드 보관 권장
SVG를 PowerPoint 도형으로 분해하면 `Alt Text` / `Shape.Tags` 가 지워져 더블클릭으로 Ketcher를 다시 열 수 없게 됩니다. 변환 전 원본 KET 도형을 **숨긴 백업 슬라이드**에 한 장 남겨 두세요.
:::

* **관련 가이드:** [**4. A0 포스터(SVG) · Q&A · 발표장 공유**](/ko/powerpoint/publishing-qa) · [**스타일 & 출력 상세 매뉴얼**](/ko/editor/publishing)

---

## 다음 단계

- [**8슬라이드 PowerPoint Playbook (.pptx) 실습하기**](/ko/powerpoint/playbook) — 발표자 노트를 켠 채 텍스트 상자 `0.1초` 변환, 4카드 그리드, KET 반응식, `1CRN` 모핑을 직접 체험해 보세요.
- [**10초 원클릭 설치 가이드**](/ko/guide/install) — PowerPoint·Word·Excel 통합 설치 명령어를 확인하세요.
