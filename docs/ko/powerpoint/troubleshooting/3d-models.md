# 3D 모델 · 편집

PowerPoint에서 360° 회전 3D `.glb` 단백질 모델, 슬라이드 더블클릭 재편집, 발표장 PC 호환성 문제 해결 방법입니다.

## 1. 3D `.glb` 모델이 삽입되지 않거나 회전하지 않음

* **지원 버전 요구사항:** 대화형 3D `.glb` 모델은 **PowerPoint 2019, 2021, 2024 또는 Microsoft 365**에서 지원됩니다.
* **구버전 PowerPoint (2016) 또는 정적 PDF 배포본:** Mol\* 상단 툴바에서 **`3D Model (GLB)`** 대신 **`2D Picture`** 모드를 선택하면 모든 PowerPoint 버전에서 선명하게 표시되는 고해상도 PNG 그림으로 삽입됩니다.

## 2. 슬라이드에서 더블클릭해도 Ketcher / Mol\*이 열리지 않음

`func_SMILES2IMG`는 각 도형의 **대체 텍스트(Alt Text)**(`SMILES:...`, `KET:...`, `MOLSTAR:...`)와 `CustomXMLParts` 메타데이터로 2D 캔버스·3D 카메라 각도를 복원합니다.

* **일반 스크린샷:** 화학 메타데이터가 없는 캡처 이미지는 PowerPoint 기본 그림 서식 창이 열립니다.
* **대체 텍스트 확인:** 도형 우클릭 ➔ **대체 텍스트 보기**에서 태그가 유지되고 있는지 확인하세요.
* **단축키 대안:** 도형 선택 후 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> (2D) 또는 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** (3D)를 클릭하세요.
* **COM 추가 기능 확인:** `[SMILES]` 탭 자체가 보이지 않으면 [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)를 확인하세요.

## 3. 추가 기능이 없는 발표장 PC · Mac에서 발표하기

삽입된 2D PNG와 3D `.glb` 모델은 PowerPoint 표준 개체이므로, `func_SMILES2IMG`가 없는 학회장 PC나 **macOS**에서도 3D 회전을 포함해 그대로 발표할 수 있습니다. 자세한 내용은 [문서 공유](/ko/guide/troubleshooting/document-sharing)를 참고하세요.

## 관련 가이드

- [**다크 슬라이드 배경**](/ko/powerpoint/troubleshooting/dark-slides) — 하얀 네모 배경 제거 및 프레젠테이션 전체 다크 모드 변환
- [**오프라인 PDB 파일 (Mol\*)**](/ko/editor/troubleshooting/offline-pdb) — 방화벽으로 RCSB PDB 조회가 막혔을 때 로컬 .pdb / .cif 파일 열기
