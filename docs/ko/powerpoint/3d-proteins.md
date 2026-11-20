# 3D 단백질 슬라이드 (`Mol*`)

내장 **Mol\*** 3D 뷰어로 PowerPoint 슬라이드 위에 **마우스로 360° 돌려볼 수 있는 3D 단백질 모델(`.glb`)** 또는 투명 배경 고해상도 2D 그림을 직접 삽입합니다.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

## 사용 순서

1. PowerPoint 상단 리본에서 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 버튼을 클릭합니다.
2. 4자리 PDB 코드(`1CRN`, `4HHB`, `6LU7`, `1BNA`)를 입력해 **`Fetch`** 버튼을 누르거나 **`Open File...`** 버튼으로 로컬 `.pdb` / `.cif` / `.sdf` 파일을 엽니다.
3. 출력 모드를 선택하고 **`Apply to Slide`** (<kbd>Ctrl</kbd> + <kbd>Enter</kbd>)를 누릅니다:
   - **`3D Model (GLB)`** — 슬라이드에서 마우스로 360° 회전 시연 (PowerPoint 2019+/365 지원, 추가 기능이 없는 Mac/학회장 PC에서도 회전 가능)
   - **`2D Picture`** — 고해상도 PNG 렌더링 (다크 슬라이드에서는 `BG: Transparent` 권장)
4. 삽입된 3D 모델·2D 그림을 **더블클릭**하면 저장된 카메라 각도와 스타일 그대로 Mol\*이 다시 열립니다.
