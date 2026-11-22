# 다크 슬라이드 · 그룹 도형 문제 해결

PowerPoint 다크 테마 슬라이드에 하얀 사각형 배경이 남거나 검은 선이 안 보일 때, 그룹화된 카드 도형이 `Batch Style` 에서 건너뛰어질 때, 또는 다크 발표 자료를 밝은 템플릿·흑백 유인물로 되돌릴 때의 해결 가이드입니다.

## 1. 다크 슬라이드에 하얀 네모 배경이 생기거나 검은 선이 안 보임 {#dark-slides}

일반 캡처 그림을 어두운 남색/슬레이트 슬라이드에 올리면 눈부신 흰 사각형 배경이 남고, 배경만 투명한 검은 선화는 어두운 바탕에 묻혀 안 보이게 됩니다.

* **기존 도형 · 발표 덱 전체 1클릭 해결:** 해당 구조식을 선택하거나 **빈 슬라이드 배경(선택 없음)** 을 클릭한 상태에서 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`** 를 실행하면 슬라이드 전체가 투명 흰색 선화로 바뀝니다.
* **처음부터 투명하게 넣기:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd> 를 누르기 전 Ketcher 상단 바를 **`[Background: Transparent]`** 로, Mol\* 상단 바를 **`BG: Transparent`** 로 설정해 두세요.

## 2. 그룹화(`Ctrl+G`)한 카드 안의 구조식이 `Batch Style` 에서 빠짐 {#grouped-shapes}

* **원인:** `Batch Style` 은 슬라이드의 최상위 도형 목록(`slide.Shapes`)만 순회하며, 카드 텍스트 상자와 묶인 **그룹 도형(`Ctrl+G`)** 내부 구조는 레이아웃 흐트러짐을 막기 위해 의도적으로 건너뜁니다.
* **해결 방법:** 그룹화된 후보 카드를 선택 ➔ <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>G</kbd> (**그룹 해제**) ➔ **`[SMILES]` ➔ `[Batch Style]`** 실행 후 필요 시 다시 <kbd>Ctrl</kbd> + <kbd>G</kbd> 로 묶으세요.

## 3. 다크 발표 슬라이드를 밝은 랩미팅 · 흑백 유인물로 되돌리기

학회용 다크 모드 발표 자료를 밝은 랩미팅 템플릿이나 배포용 흑백 유인물로 재활용할 때도 구조식을 다시 그릴 필요가 없습니다:

* **밝은 템플릿용 컬러 복원:** 빈 슬라이드 여백 클릭 ➔ **`[SMILES]` ➔ `[Batch Style]` ➔ `[Standard Color (Transparent)]`** 로 투명 배경 IUPAC 컬러 선화로 즉시 복구됩니다.
* **인쇄용 흑백 유인물 변환:** `.pptx` 사본을 저장한 뒤 **`[Batch Style]` ➔ `[Publication B&W (ACS Monochrome)]`** 를 실행해 1초 만에 순수 흑백으로 바꾼 뒤 PDF로 내보내세요.

::: tip 애니메이션이 걸린 도형 주의사항
도형에 PowerPoint 나타나기·이동 애니메이션이 적용되어 있다면 화학 편집과 `Batch Style` 통일을 먼저 끝낸 뒤 마지막에 애니메이션을 거는 것이 편리합니다.
:::

## 관련 가이드

- [**3D 모델 · 더블클릭 편집**](/ko/powerpoint/troubleshooting/3d-editing) — `.glb` 용량·버전 제한, 더블클릭 미작동, 학회장 PC/Mac 호환 점검
- [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) — 투명 4-카드 비교 그리드 제작 및 `Batch Style` 4종 프리셋
- [**리본 · 우클릭 메뉴**](/ko/powerpoint/reference) — PowerPoint `[SMILES]` 리본 전체 명령표 및 5종 우클릭 메뉴
