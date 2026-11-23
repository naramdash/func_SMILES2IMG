# 3. 3D 단백질 Figure (`Mol*`)

내장 **Mol\*** 3D 뷰어로 Microsoft Word 논문과 보고서 안에 단백질, 핵산, 리간드 결합 포켓 Figure를 삽입합니다. 4자리 PDB 코드 조회는 물론 로컬 `.pdb` / `.cif` 파일을 오프라인으로 열 수 있고, 저장된 카메라 각도 그대로 더블클릭 재편집이 가능합니다.

<FeatureMedia id="word-ch5-edit-and-molstar" />

## 1. 논문 본문에 3D 단백질 Figure 넣기

1. Word 본문에서 그림을 배치할 위치에 커서를 둡니다.
   ::: tip 숨은 꿀팁: 본문 속 PDB 코드를 드래그해 바로 열기
   원고 본문에 `1CRN`, `6LU7` 같은 4자리 PDB 코드나 로컬 `.pdb`/`.cif` 경로가 적혀 있다면 **텍스트를 드래그 선택한 상태에서 `[3D Structure (Mol*)]`를 클릭**하세요. Mol\* 창이 열리자마자 해당 구조가 자동으로 로드됩니다.
   :::
2. 상단 리본에서 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `3D Structure (Mol*)...`**)를 클릭합니다.
3. 두 가지 방식 중 하나로 구조를 불러옵니다:
   * **공개 PDB 코드 조회:** 4자리 PDB ID(`1CRN`, `4HHB`, `6LU7`, `1BNA`) 입력 후 **`Fetch`** 클릭.
   * **오프라인 로컬 파일 열기:** **`Open File...`** 로 내 PC의 `.pdb`, `.cif`, `.bcif`, `.ent`, `.mol`, `.sdf`, `.molj`, `.molx` 파일을 외부 전송 없이 엽니다.
4. **Cartoon**, **Surface**, **Ball & Stick**(리간드 결합 포켓) 스타일과 배경(`White` / `Transparent`)을 맞춘 뒤, 상단 툴바에서 출력 모드를 골라 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**)를 누릅니다.

## 2. `2D Picture` vs. `3D Model (GLB)` 선택 기준

| Word 내 Mol\* 출력 모드 | `.docx`에서의 동작 특징 | 추천 용도 |
| :--- | :--- | :--- |
| **`2D Picture` (논문 투고 권장)** | 고해상도 `.png` 이미지로 삽입 (`BG: White` 또는 `Transparent`) | **학술지 투고, PDF 변환 제출, 인쇄용 논문 Figure** (모든 Word 버전 호환) |
| **`3D Model (GLB)`** | Word 페이지 위에서 마우스로 360° 회전할 수 있는 대화형 3D 개체 삽입 | **사내 기술 보고서, 연구비 과제 검토서, 전자 문서 열람용** (Word 2019 / 2021 / 2024 / 365 지원) |

::: info 배경색 & 멀티패널 Figure 조합 팁
학술지 흰색 지면과 PDF 제출물에는 **`BG: White`** 또는 **`BG: Transparent`** 가 가장 깔끔합니다. 도형이나 음영 박스 위에 겹쳐 배치해야 할 때는 **`BG: Transparent`** 를 선택해 흰 여백을 없애세요.
:::

## 3. 카메라 앵글 보존 재편집 & 스타일 보호

* **정확한 뷰어 상태 저장:** `func_SMILES2IMG`는 삽입 시 Mol\* 세션 상태(PDB ID, 카메라 각도, 줌 배율, 표현 스타일, 배색)를 압축해 Word 문서 내부 **`CustomXMLParts`** 에 저장하고 그림의 Alt Text(`MOLSTAR:...`)와 연결합니다.
* **더블클릭 제자리 수정:** 삽입된 2D 단백질 그림이나 3D 모델을 **더블클릭**하면 마지막에 맞춘 각도 그대로 Mol\*이 다시 열려 세부 각도나 잔기 표시만 살짝 다듬어 교체할 수 있습니다.
* **`Batch Style` 안전 제외:** 투고 전 문서 전체 2D 구조식을 **`[Batch Style]` ➔ `[Publication B&W]`** 로 흑백 변환하더라도 Mol\* 3D 그림은 건드리지 않으므로 단백질 컬러 렌더링이 그대로 보존됩니다.

## 관련 가이드

- [**Mol\* 3D 뷰어 상세 매뉴얼**](/ko/editor/molstar) — 추천 PDB 코드 4선(`1CRN`, `4HHB`, `6LU7`, `1BNA`) 및 결합 포켓 스타일링 가이드
- [**4. 투고 마감 · 벡터 SVG · 공동저자 공유**](/ko/word/publishing) — 2D 구조식만 일괄 흑백 마감하고 공동저자와 `.docx` 공유하기
- [**오프라인 PDB 파일 트러블슈팅**](/ko/editor/troubleshooting/offline-pdb) — 사내 방화벽 환경에서 로컬 `.pdb` / `.cif` 파일 활용법
