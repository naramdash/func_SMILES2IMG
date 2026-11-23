# 1. 2D 구조식 · 반응식 슬라이드

슬라이드에 적어 둔 텍스트 상자(SMILES·InChI·화합물명)를 그 위치에서 구조식 도형으로 변환하거나, **Ketcher 2D**로 다단계 합성 반응 메커니즘 슬라이드를 그리고 언제든 **더블클릭**해 카드 정렬을 유지한 채 제자리 수정합니다.

## 1. 슬라이드에 2D 구조식을 넣는 방법

| 현재 슬라이드 상태 | 실행할 동작 | PowerPoint 배치 결과 |
| :--- | :--- | :--- |
| **텍스트 상자 선택 상태** (SMILES, InChI, KET 또는 화합물명/CAS `Aspirin`, `Paracetamol`, `50-78-2`) | 상자 테두리를 클릭해 선택한 뒤 **`[SMILES]` ➔ `[Paste Structure]`** | **해당 텍스트 상자 자리에 제자리 치환**되며 비율 잠긴 구조식으로 교체 |
| **문장 일부 텍스트 드래그** 또는 **클립보드 복사 상태** | 글자 일부 드래그(또는 아무것도 선택 안 함) ➔ **`[Paste Structure]`** | 기존 텍스트 상자를 지우지 않고 **슬라이드 중앙**에 새 구조식 삽입 |
| **적어 둔 물질명/SMILES를 살짝 고쳐 넣고 싶을 때** | 텍스트 상자나 글자를 선택한 뒤 **`[Insert / Edit Structure]`** | **Ketcher 2D 캔버스에 해당 구조가 미리 그려진 채 열려** 바로 수정 후 삽입 가능 |
| **빈 슬라이드** (새로 그리기) | **`[SMILES]` ➔ `[Insert / Edit Structure]`** | 오프라인 **Ketcher 2D** 열림 ➔ <kbd>Ctrl</kbd> + <kbd>Enter</kbd> 시 슬라이드 중앙 삽입 |

::: tip 텍스트 상자 제자리 변환 (`Paste Structure`)
<FeatureMedia id="powerpoint-slide-textbox-quick-insert" />

`CC(=O)Oc1ccccc1C(=O)O`나 `Paracetamol`이 들어 있는 텍스트 상자 테두리를 선택하고 **`[Paste Structure]`** 를 누르면 그 텍스트 상자 위치와 중심에 맞춰 깔끔한 크기·고정 비율의 구조식으로 즉시 교체됩니다. *(한 줄짜리 얇은 텍스트 상자라도 보기 좋은 최소 높이를 확보해 슬라이드 안쪽에 예쁘게 배치합니다.)*
:::

## 2. 다단계 합성 반응 메커니즘 슬라이드 (`KET`)

반응 화살표(`➔`), `+` 기호, 촉매·온도·수율 주석이 포함된 유기 합성 반응 경로를 `.pptx` 안에 무손실로 저장합니다:

<FeatureMedia id="powerpoint-slide-reaction" />

1. 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** 를 클릭합니다 (단축키는 [Ketcher 2D 상세 매뉴얼](/ko/editor/) 참고).
2. **Ketcher 2D**에서 반응물, 화살표(`➔`), `+` 기호, 시약·촉매 텍스트(`T` 도구)를 그립니다.
3. 어두운 슬라이드나 카드 위에 올릴 때는 상단 바를 **`[Background: Transparent]`** 로 맞춘 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Slide`**)를 누릅니다. *(흰 배경으로 넣었더라도 나중에 `[Batch Style] ➔ [Dark Mode]`로 한 번에 바꿀 수 있습니다.)*
4. **`.pptx` 무손실 보존:** 반응식 전체 KET JSON 좌표가 도형의 `Shape.Tags["KET"]`와 `AlternativeText`에, 표준 SMILES가 `Shape.Tags["SMILES"]`와 `Title`에 함께 저장되어 저장·슬라이드 복사·재오픈 후에도 화살표와 텍스트 배치가 그대로 유지됩니다.

## 3. 더블클릭 제자리 수정 · 애니메이션 팁 · 크로스 앱 재활용

* **카드 위치·크기 유지:** 일반 편집 화면(`Slide View`)에서 슬라이드 편집 캔버스 위의 2D 구조식이나 반응식 도형을 **더블클릭**하면 Ketcher 2D가 바로 열립니다. 리본이나 다른 작업 창 더블클릭은 가로채지 않으며, 적용 시 기존 도형 박스(`Left`, `Top`, `Width`, `Height`) 안에 맞추어 교체되므로 카드 정렬이 흐트러지지 않습니다.
* **물질명/CAS 안전 확인 & 캐시:** SMILES·InChI·KET·검증된 CAS 번호는 즉시 변환되며, 일반 영문 물질명이나 모호한 짧은 문자열은 **`Paste Structure`** 실행 시 변환 전 확인 창이 표시됩니다. 사전에 없는 물질명 조회 결과는 로컬 디스크에 **30일간** 캐시되며 잘못된 체크 디지트 CAS는 즉시 거부됩니다.
* **애니메이션 적용 순서 팁:** 재편집 시 기존 도형 개체를 교체(`Delete` + `AddPicture`)하므로, **나타내기/페이드 애니메이션은 화학 구조 수정을 마친 마지막 단계에 추가**하는 것이 좋습니다.
* **Word 논문 · Excel 시트에서 발표 슬라이드로 가져오기:**
  * **Word ➔ PowerPoint:** 논문 원고의 반응식을 클릭해 **`[Copy KET]`**(단일 분자는 **`[Copy SMILES]`**) 복사 ➔ PowerPoint에서 **`[Paste Structure]`** ➔ **`[Batch Style]` ➔ `[Dark Mode]`** 실행.
  * **Excel ➔ PowerPoint:** Excel SAR 시트 셀에서 **`[Copy SMILES]`** / **`[Copy KET]`** 복사 ➔ PowerPoint에서 **`[Paste Structure]`** 실행.

## 관련 가이드

- [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) — 투명 흰색 선화 배치, 4-카드 그리드 및 덱 전체 테마 일괄 전환
- [**Ketcher 2D 편집기 상세 매뉴얼**](/ko/editor/) — 고리 템플릿, 단축키, 배경 옵션 전체 안내
- [**3D 모델 · 더블클릭 문제 해결**](/ko/powerpoint/troubleshooting/3d-editing#double-click) — 더블클릭이 안 열릴 때 확인 사항
