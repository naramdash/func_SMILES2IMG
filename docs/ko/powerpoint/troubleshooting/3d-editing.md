# 3D 모델 · 더블클릭 · 발표장 호환성

PowerPoint에서 Mol\* `.glb` 3D 단백질 삽입·회전 오류와 `.pptx` 용량 문제, 슬라이드 도형 더블클릭 재편집 미작동, 학회장 연단 PC 및 macOS 발표 호환성 트러블슈팅입니다.

## 1. 3D `.glb` 모델이 삽입·회전되지 않거나 `.pptx` 용량이 너무 큼 {#3d-models}

* **지원 Office 버전 확인:** PowerPoint 내장 3D `.glb` 마우스 회전과 **모핑(Morph)** 카메라 전환은 **PowerPoint 2019, 2021, 2024, Microsoft 365** 에서 지원됩니다.
* **PowerPoint 2016 · 저용량 배포본 · 초대형 리보솜/캡시드:** Mol\* 상단 툴바를 **`3D Model (GLB)`** 에서 **`2D Picture`** (`BG: Transparent`) 로 바꾸면 용량이 작고 모든 PowerPoint 버전에서 선명하게 표시되는 투명 PNG가 들어갑니다.
* **기관 방화벽으로 PDB 다운로드가 막힌 경우:** Mol\* 창에서 **`Open File...`** 을 눌러 내 PC의 `.pdb`, `.cif`, `.bcif`, `.ent`, `.mol`, `.sdf`, `.molj`, `.molx` 파일을 100% 오프라인으로 여세요 ([오프라인 PDB 파일](/ko/editor/troubleshooting/offline-pdb)).

## 2. 슬라이드에서 더블클릭해도 Ketcher / Mol\* 이 열리지 않음 {#double-click}

`func_SMILES2IMG` 는 **`Shape.Tags`** (`KET`, `SMILES`, `MOLSTAR`) 와 **대체 텍스트(Alt Text)**, 그리고 3D 세션용 `CustomXMLParts` 로 편집 가능한 슬라이드 도형을 식별합니다.

* **일반 캡처 이미지 · 캔버스 밖 클릭:** 외부 스크린샷이나 PDF 복사 이미지에는 화학 태그가 없어 PowerPoint 기본 그림 서식 창이 열립니다. **`[Insert / Edit Structure]`** 나 **`[Paste Structure]`** 로 다시 삽입하세요. *(참고: 더블클릭 편집은 일반 편집 보기의 슬라이드 캔버스 영역에서만 동작합니다.)*
* **도형으로 변환한 SVG:** **`[Copy SVG]`** 로 붙여넣고 `그룹 해제(도형으로 변환)` 한 벡터 다각형은 더 이상 Ketcher에 연결되어 있지 않습니다. 원본 KET 도형을 숨김 백업 슬라이드에 남겨 두세요.
* **태그 확인 · 리본 메뉴 실행:** 도형 **우클릭 ➔ 대체 텍스트 보기**에서 `SMILES:`, `KET:`, `MOLSTAR:` 접두어가 남아 있는지 확인하세요. 언제든 도형을 선택한 뒤 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) 또는 **`[3D Structure (Mol*)]`** (3D) 로 열 수 있습니다. `[SMILES]` 탭이 안 보이면 [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)를 점검하세요.

## 3. 추가 기능 없는 발표장 PC · Mac 에서 발표하기 {#sharing}

* **발표장 PC · macOS 무설치 재생:** 삽입된 2D PNG와 3D `.glb` 모델은 `.pptx` 안에 내장되는 PowerPoint 표준 개체이므로, `func_SMILES2IMG` 가 없는 **학회장 연단 PC나 macOS** 에서도 3D 회전·모핑 애니메이션이 그대로 작동합니다 ([문서 공유 가이드](/ko/guide/troubleshooting/document-sharing)).
* **PDF 유인물 저장 팁:** `.pptx` 를 PDF로 내보내면 3D `.glb` 모델은 현재 정지한 각도로 고정되므로, 저장 전 원하는 결합 포켓 각도가 맞춰져 있는지 확인하세요.

## 관련 가이드

- [**다크 슬라이드 · 그룹 도형**](/ko/powerpoint/troubleshooting/dark-slides) — 하얀 네모 배경, 어두운 선화, 그룹화(`Ctrl+G`) 누락 해결
- [**3. 360° 3D 단백질 발표 (`Mol*` & 모핑)**](/ko/powerpoint/3d-proteins) — 라이브 3D 모핑 줌인 연출 및 경량 `.glb` 활용법
- [**오프라인 PDB 파일**](/ko/editor/troubleshooting/offline-pdb) · [**문서 공유 가이드**](/ko/guide/troubleshooting/document-sharing) — 방화벽 환경 로컬 파일 열기 및 학회장 호환성 안내
