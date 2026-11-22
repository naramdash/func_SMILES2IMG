# 일괄 스타일 · 17열 TSV · 3D (`Mol*`)

클릭 한 번으로 선택 범위 또는 활성 엑셀 시트 전체의 구조식 수식을 일괄 재작성하고, 분자 1개를 인접 17개 열의 물성 데이터(TSV)로 바로 펼치며, 셀에 적힌 4자리 PDB 코드를 자동 인식해 360° 회전 3D 단백질(`Mol*`)을 시트 위에 배치합니다.

## 1. 수식 자동 재작성 일괄 스타일 (`Batch Style`)

상단 **`[SMILES]`** ➔ **`[Batch Style]`** 을 클릭하면 선택한 범위나 시트 전체 구조식의 배경과 결합선 색상을 한 번에 통일할 수 있습니다:

* **수식 인수 직접 갱신:** 셀에 들어 있는 모든 `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG` 및 `.FLOAT` 수식에서 첫 번째 입력 인수(`A2`, `"Aspirin"` 등)는 그대로 두고 `[bg]` 와 `[style]` 인수를 새 프리셋으로 자동 재작성합니다 (기존에 수동 지정한 스타일/변환 인수는 새 프리셋으로 일괄 대체됩니다).
* **독립 2D 그림 도형 동시 반영:** 여러 셀을 드래그하지 않은 경우 활성 시트 위의 일반 리본 삽입 2D 그림 도형도 함께 스타일이 업데이트됩니다 (`Mol*` 3D 개체와 수식으로 관리되는 `.FLOAT` 도형은 중복 처리 없이 적절히 구분됩니다).
* **적용 범위 선택:** 특정 표 범위만 변환하려면 2개 이상의 셀을 드래그하고, 단일 셀만 선택한 상태에서는 **현재 활성 시트 전체**(`UsedRange` 최대 500행 × 50열)가, 2D 그림 도형을 선택한 상태에서는 해당 도형만 변환됩니다.

| `Batch Style` 프리셋 | 자동 기록되는 Excel 수식 | 추천 시트 활용처 |
| :--- | :--- | :--- |
| **`Standard Color (Transparent)`** | `=SMILES2IMG(A2, "trans", "color")` | 투명 배경 + IUPAC 표준 원소 컬러 기본 표 |
| **`Publication B&W (ACS Monochrome)`** | `=SMILES2IMG(A2, "trans", "bw")` | 학술지 투고 요약표, 특허 명세서, 흑백 인쇄 |
| **`White Background (Color)`** | `=SMILES2IMG(A2, "white", "color")` | 음영·조건부 서식이 칠해진 행 위 구조식 분리 |
| **`Dark Mode (White Ink)`** | `=SMILES2IMG(A2, "trans", "white")` | 어두운 배경의 대시보드 시트 위 흰색 결합선 |

## 2. 17열 TSV 행 확장 & 수식 식별자 역추출

수식 구조식 셀(`=NAME2IMG`, `=SMILES2IMG`, `=INCHI2IMG`, `=KET2IMG`) 또는 삽입된 그림 도형을 선택한 상태에서 **`[SMILES]`** 리본(또는 우클릭 ➔ **`Chemical Structure`**)으로 물성 표를 확장하거나 원본 데이터를 추출합니다:

<FeatureMedia id="common-molecule-info" />

* **`Copy as Excel Row (TSV Columns)` (`Molecule Info`):** `Name`, `Formula`, `MW`, `Exact Mass`, `LogP`, `TPSA`, `HBD`, `HBA`, `RotBonds`, `Heavy Atoms`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski Ro5`, `Canonical SMILES`, `InChIKey`까지 **17개 지표를 탭으로 구분해 복사**합니다. 수식 17개를 일일이 쓰지 않고 <kbd>Ctrl</kbd> + <kbd>V</kbd> 한 번으로 옆 열(`A:Q`)에 고정값으로 바로 채울 때 유용합니다 ([5종 복사 서식 규격 →](/ko/editor/publishing)).
* **`Copy SMILES` / `Copy InChI` / `Copy KET`:** `=NAME2IMG("Aspirin")` 이나 CAS 번호 수식 셀, `.FLOAT` 도형에서도 해석된 표준 SMILES·InChI·KET JSON을 즉시 추출합니다.
* **`Copy SVG`:** 벡터 그래픽(`image/svg+xml` 및 `<svg>` XML)으로 복사해 스크리닝 표의 선도 물질을 Adobe Illustrator · Inkscape로 깨짐 없이 옮깁니다.

## 3. 셀 PDB 코드 연동 3D 단백질 (`Mol*`)

내장 **Mol\*** 3D 뷰어를 통해 엑셀 시트 위에 마우스로 360° 회전하는 3D 단백질(`.glb`) 또는 투명 2D 그림(`.png`)을 바로 올릴 수 있습니다:

<FeatureMedia id="common-molstar-3d" />

1. **셀의 4자리 PDB 코드 자동 로드:** 엑셀 셀에 `1CRN`, `4HHB`, `6LU7`, `1BNA` 같은 4자리 PDB 코드가 입력된 상태에서 셀을 선택하고 **`[SMILES]`** ➔ **`[3D Structure (Mol*)]`** 를 클릭하면 해당 단백질이 자동으로 열립니다. *(빈 셀에서 열면 PDB/AlphaFold 검색이나 로컬 `.pdb` · `.cif` · `.sdf` 파일 오프라인 열기가 가능합니다.)*
2. **표현 스타일 및 출력 형식 선택:** **Cartoon**, **Surface**, **Ball & Stick** 등을 설정하고 **`3D Model (GLB)`** (시트 위 3D 회전 개체) 또는 **`2D Picture`** (고해상도 투명 PNG)를 고릅니다.
3. **동일 각도로 언제든 재편집:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd>로 시트에 배치한 뒤, 나중에 단백질 개체를 **더블클릭**하면 마지막 카메라 시점과 표현 그대로 Mol\*이 다시 열립니다.

## 관련 가이드

- [**리본 · 우클릭 메뉴**](/ko/excel/ribbon) — Excel `[SMILES]` 리본 전체 명령, 3종 우클릭 메뉴 및 Indigo 파서 진단 로그
- [**SDF 파일 입출력**](/ko/excel/sdf) — `.sdf`/`.mol` 파일을 12열 라이브 시트로 가져오거나 V2000 `.sdf` 로 내보내기
- [**Mol\* 3D 뷰어 상세 매뉴얼**](/ko/editor/molstar) — 단백질 표현 방식, 리간드 포켓 강조 및 오프라인 로컬 PDB 파일 활용
