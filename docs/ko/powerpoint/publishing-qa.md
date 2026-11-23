# 4. A0 포스터(SVG) · Q&A · 발표장 공유

학회 포스터 인쇄 및 구두 발표 전 점검 사항을 정리한 **발표 · 포스터 체크리스트**입니다.

<FeatureMedia id="powerpoint-slide-live-qa-info" />

## 발표 · 포스터 체크리스트

### Step 1. 대형 A0 · A1 학회 포스터용 무손실 벡터 `SVG` 복사 (`Copy SVG`)

일반 PNG/JPG 캡처는 노트북 화면에서는 괜찮아 보여도, 가로·세로 1미터 크기의 **A0 / A1 학회 포스터(90 × 120 cm, 300 DPI)** 로 출력하면 결합선과 원소 기호가 흐려지거나 픽셀이 깨집니다.

1. 슬라이드의 분자 구조식·반응 메커니즘을 선택합니다(또는 텍스트 상자 속 SMILES/화합물명 문자열 드래그).
2. 상단 **`[SMILES]` ➔ `[Copy SVG]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Copy Vector SVG`**)를 클릭합니다.
3. PowerPoint 포스터 슬라이드(Office 365/2019+)나 **Adobe Illustrator / Inkscape**에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣으면 확대 배율과 무관하게 매끈한 벡터 곡선이 유지됩니다.

### Step 2. 발표 Q&A 물성 답변 준비하기 (`Molecule Info`)

PowerPoint 전체 화면 **슬라이드 쇼 진행 중에는 상단 리본과 추가 기능 우클릭 메뉴가 표시되지 않습니다.** 발표 상황에서는 아래 두 단계로 대비하는 것이 좋습니다:

* **발표 전 준비 (발표자 노트 활용):** 핵심 후보 물질 도형을 선택해 **`[SMILES]` ➔ `[Molecule Info]` ➔ `[Copy as Multi-Line List]`** 를 누르고, 화학식·HRMS 정밀질량·LogP·TPSA·`Fsp3` 정보를 **발표자 노트**에 붙여넣어 두세요. 발표자 보기 화면에서 청중 질문 시 바로 읽어 답할 수 있습니다.
* **리허설 · 랩미팅 실시간 조회:** 심층 Q&A가 들어오면 <kbd>Esc</kbd>로 슬라이드 쇼를 잠시 빠져나와 해당 구조식(또는 SMILES 텍스트)을 클릭하고 **`[Molecule Info]` ➔ `[Detailed Molecule Info Window...]`** 를 띄워 12종 물성 지표를 보여 주세요(행 더블클릭으로 개별 값 복사 가능, 5가지 복사 서식 전체 예시는 [스타일 · 출력 상세 매뉴얼](/ko/editor/publishing) 참고).

### Step 3. 발표장 연단 PC · Mac · PDF 유인물 최종 점검

* **학회장 PC · macOS 호환:** 삽입된 2D PNG와 3D `.glb` 모델(모핑 포함)은 PowerPoint 표준 개체이므로 `func_SMILES2IMG`가 설치되지 않은 **발표장 Windows PC와 macOS**에서도 동일하게 재생됩니다.
* **PDF 유인물 · 흑백 배포본 생성:**
  * `.pptx`를 PDF로 저장하면 3D `.glb` 모델은 현재 정지된 카메라 각도로 고정되므로, 저장 전 각 3D 슬라이드의 기본 각도가 잘 잡혀 있는지 확인하세요.
  * 배포용 흑백 유인물이 필요할 때는 사본을 저장한 뒤 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`** 로 일괄 변환해 PDF로 내보내세요.

## 관련 가이드

- [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) — 발표 덱 전체 테마를 다크 모드·컬러·흑백으로 일괄 전환
- [**3. 360° 3D 단백질 발표 (`Mol*` & 모핑)**](/ko/powerpoint/3d-proteins) — `.glb` 3D 단백질 삽입 및 모핑 줌인 연출
- [**리본 · 우클릭 메뉴**](/ko/powerpoint/reference) · [**트러블슈팅**](/ko/powerpoint/troubleshooting/dark-slides) — 리본·우클릭 치트시트 및 다크 배경·그룹 도형·3D 모델 문제 해결
