# 2. 다크 모드 · 후보 물질 카드 · 덱 테마 전환

어두운 네이비·슬레이트·그라데이션 PowerPoint 발표 템플릿 위에 하얀 네모 여백 없이 투명 흰색 선화 구조식을 배치하고, 4-카드 후보 물질 비교 그리드 제작부터 **프레젠테이션 전체 테마 1초 전환**까지 한 번에 처리합니다.

## 1. 다크 슬라이드에서 하얀 네모 배경 없애는 2가지 방법

일반 스크린샷을 어두운 슬라이드에 올리면 눈부신 흰색 사각형 배경이 남습니다. `func_SMILES2IMG`는 완전한 알파 채널 투명 배경과 고대비 흰색 결합선을 제공합니다:

<FeatureMedia id="powerpoint-slide-dark-mode" />

* **방법 A — 삽입 전 투명 설정 (Ketcher / Mol\*):** <kbd>Ctrl</kbd> + <kbd>Enter</kbd> 전에 Ketcher 상단 바를 **`[Background: Transparent]`** 로, Mol\* 상단 바를 **`BG: Transparent`** 로 선택합니다.
* **방법 B — 삽입 후 클릭 한 번 변환 (`Batch Style`):** 구조식 도형을 선택하고(아무것도 선택하지 않으면 **덱 전체**) **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`** 를 클릭합니다.

## 2. 4-카드 후보 물질 SAR 비교 그리드 만들기

신약 후보 물질(Hit-to-Lead) 비교 슬라이드에서 구조식·활성 지표·물성 요약을 정갈하게 정렬하는 레시피입니다:

<FeatureMedia id="powerpoint-slide-candidate-grid" />

1. 슬라이드 배경을 짙은 슬레이트/네이비(`#0F172A`, `#1E293B`)로 두고, 동일한 크기의 둥근 사각형 카드 4개를 나란히 배치합니다.
2. 각 카드 자리에 **`[Paste Structure]`**(텍스트 상자 제자리 변환) 또는 **`[Insert / Edit Structure]`** 로 구조식을 넣습니다.
3. **`[SMILES]` ➔ `[Batch Style]` ➔ `[Dark Mode (White Ink)]`** 를 눌러 4개 카드의 모든 구조식을 균일한 굵기의 투명 흰색 선화로 통일합니다.
4. 각 구조식을 선택한 채 **`[SMILES]` ➔ `[Molecule Info]` ➔ `[Copy as Multi-Line List]`** (또는 **`Copy Molecular Weight`**)를 누르면 화학식·분자량·LogP·TPSA·`Fsp3` 지표를 카드 하단 텍스트 상자에 오타 없이 붙여넣을 수 있습니다.
5. **정렬 흐트러짐 없는 치환기 수정:** 카드 위 구조식을 **더블클릭**해 R-그룹을 바꿔도 카드 중앙 위치와 크기가 그대로 유지됩니다.

## 3. 발표 덱 전체 테마 1초 일괄 전환 (`Batch Style`)

학회용 다크 모드 발표 자료를 밝은 랩미팅 템플릿이나 인쇄용 흑백 유인물로 재활용할 때 슬라이드를 하나씩 고칠 필요가 없습니다:

<FeatureMedia id="powerpoint-slide-batch-style" />

1. **적용 범위 지정:**
   * **프레젠테이션 전체 슬라이드 변환:** 빈 슬라이드 배경을 클릭해 아무 도형도 선택되지 않은 상태에서 실행하면 **모든 슬라이드의 2D 구조식·반응식**이 한 번에 바뀝니다 (`Mol*` 3D 개체는 유지).
   * **특정 도형만 변환:** <kbd>Shift</kbd> + 클릭으로 원하는 구조식 도형만 골라 실행합니다.
2. 상단 **`[SMILES]` ➔ `[Batch Style]`** 에서 원하는 테마를 고릅니다 (전체 규격은 [스타일 & 출력 상세 매뉴얼](/ko/editor/publishing) 참고):
   * **`Dark Mode (White Ink)`** — 투명 배경 + 흰색 결합선 (어두운 슬라이드용)
   * **`Standard Color (Transparent)`** — 투명 배경 + IUPAC 원소 컬러 (밝은/파스텔 슬라이드용)
   * **`Publication B&W (ACS Monochrome)`** — 투명 배경 + 순수 흑백 선화 (흑백 인쇄·유인물용)
   * **`White Background (Color)`** — 불투명 흰색 카드 + 표준 컬러 (사진·그라데이션 배경 위 시인성 확보)

::: warning 그룹화된 도형(`Ctrl+G`)은 그룹 해제 후 변환하세요
PowerPoint의 `Batch Style`은 슬라이드의 최상위 도형을 순회하므로, 구조식과 카드 배경을 **그룹화(Group)** 해 둔 도형은 일괄 변환에서 건너뛰게 됩니다. **`Ctrl` + `Shift` + `G`로 그룹을 해제**한 뒤 `Batch Style`을 실행하고 필요 시 다시 묶어 주세요.
:::

## 관련 가이드

- [**1. 2D 구조식 · 반응식 슬라이드**](/ko/powerpoint/slides-structures) — 텍스트 상자 0.1초 변환, Ketcher 반응식 및 더블클릭 수정
- [**4. A0 포스터(SVG) · Q&A · 발표장 공유**](/ko/powerpoint/publishing-qa) — A0 포스터 벡터 SVG 복사 및 발표 Q&A 대비
- [**다크 슬라이드 · 그룹 도형 문제 해결**](/ko/powerpoint/troubleshooting/dark-slides) — 하얀 네모 배경, 대비 부족, 그룹화 누락 해결
