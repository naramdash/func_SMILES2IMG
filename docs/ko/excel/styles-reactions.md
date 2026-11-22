# 스타일 · 부분구조 하이라이트

**`=SMILES2IMG`**, **`=NAME2IMG`**, **`=INCHI2IMG`**, **`=KET2IMG`** (및 각 `.FLOAT` 변형)까지 **8개 이미지 함수 전체가 동일한 2·3·4번째 인수 규격**을 공유합니다. 아래 키워드 한 세트로 논문용 흑백, 다크 모드, 부분구조 강조, 원자 번호, 회전을 모든 수식에 똑같이 적용할 수 있습니다.

<FormulaTable
  title="스타일 · 하이라이트 시트"
  :columns='["SMILES", "수식"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A2, \"\", \"bw\")", desc: "논문 투고용 ACS 흑백 선화 렌더링", out: "흑백 선화" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A3, \"#1E293B\", \"white\")", desc: "어두운 셀 배경 위 선명한 흰색 결합선", out: "다크 모드" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A4, \"\", \"hl:c1ccccc1\")", desc: "일치하는 벤젠 부분구조 원자 컬러 강조", out: "컬러 강조" }
  ]'
/>

## 1. 8개 이미지 함수 공용 4인수 구조

| 인수 | 위치 | 기본값 | 입력 가능한 값 |
| :--- | :---: | :--- | :--- |
| `입력` | 1번째 | *(필수)* | SMILES, 시약명/CAS, InChI, KET 셀 참조 또는 문자열 |
| `[배경색]` | 2번째 | `""` (`trans`) | `""` / `"trans"`, `"white"`, CSS 색상명 (`"yellow"`), 또는 `#RRGGBB` hex 배경 |
| `[스타일]` | 3번째 | `""` (`color`) | `"color"`, `"bw"`, `"white"`, `"ink=#HEX"`, `"hl:<smarts>"`, `"num"`, `"stereo"`, `"h"`, `"all-c"` |
| `[변형]` | 4번째 | `""` | `"flip"`/`"flipx"`, `"flipy"`, `"rot90"`, `"rot180"`, `"rot270"` *(`[스타일]` 안에 혼용 가능)* |

::: tip 흑백(`"bw"`) 2인수 단축 문법
2번째 인수 자리에 바로 `"bw"`(또는 `FALSE`)를 입력하면 빈 따옴표 `""` 없이 즉시 투명 배경 흑백 선화가 됩니다: `=SMILES2IMG(A2, "bw")`, `=NAME2IMG("Aspirin", "bw")`.
:::

## 2. 실전 스타일 · 하이라이트 · 변형 매트릭스

<FeatureMedia id="excel-style-matrix" />

| 원하는 스타일 | 수식 예시 | 시각적 효과 |
| :--- | :--- | :--- |
| **표준 IUPAC 컬러** | `=SMILES2IMG(A2)` | 원소별 색상 (`O` 빨강, `N` 파랑) + 투명 배경 기본값 |
| **논문용 ACS 흑백** | `=SMILES2IMG(A2, "", "bw")` | 학술지 투고 및 특허 명세서용 순수 흑백 선화 (`mono`, `black`, `grayscale`) |
| **불투명 흰색 카드 배경** | `=SMILES2IMG(A2, "white", "bw")` | 셀 배경색이나 눈금선 비침을 막는 순백 불투명 박스 |
| **다크 모드 (흰색 선화)** | `=SMILES2IMG(A2, "#1E293B", "white")` | 어두운 셀 배경 위 선명한 흰색 결합선 |
| **커스텀 브랜드 선 색상** | `=SMILES2IMG(A2, "", "ink=#1E40AF")` | 원하는 hex/CSS 결합선 색상 (`line=#HEX`, `base=#HEX`) |
| **부분구조 하이라이트** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | 일치하는 SMILES/SMARTS 원자·결합만 컬러로 강조 |
| **쿼리 셀 연동 하이라이트** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | 상단 셀(`$B$3`)만 바꾸면 표 전체 구조식 강조 일괄 갱신 |
| **NMR 원자 번호 · 입체** | `=SMILES2IMG(A2, "", "num stereo")` | 1번부터 원자 번호(`1, 2, 3...`) 및 CIP `R/S`, `E/Z` 표시 |
| **수소 펼침 · 탄소 `C` 표기** | `=SMILES2IMG(A2, "", "h all-c")` | 생략된 수소(`-H`)와 모든 탄소 꼭짓점(`C`) 표시 |
| **대칭 반전 · 회전** | `=SMILES2IMG(A2, "", "flip rot90")` | 좌우(`flip`)/상하(`flipy`) 반전 및 `90°`/`180°`/`270°` 회전 |

## 3. 다중 키워드 조합 · 시트 전체 일괄 변환

* **자유로운 구분자:** `[스타일]` 인수 안에서 공백(`"bw num"`), 쉼표(`"bw, num"`), 파이프(`"bw|num"`), 세미콜론(`"bw;num"`), `+`(`"bw+num"`) 중 어느 구분자도 동일하게 인식합니다.
* **작성 완료 시트 일괄 변환:** 기존 수식 수백 행을 클릭 한 번에 ACS 흑백이나 다크 모드로 바꾸려면 **`[SMILES]` ➔ `[Batch Style]`** 을 사용하세요 ([일괄 스타일 · TSV · 3D](/ko/excel/tools-export)).

## 관련 가이드

- [**40개 함수 전체 개요**](/ko/excel/formulas) — 모든 스타일 키워드 및 별칭(`idx`, `chiral`, `unfoldh`) 전체 규격
- [**부분구조 · 유사도 검색**](/ko/excel/sar) — `hl:<SMARTS>` 와 `=SMILES.MATCH` 를 결합한 12종 작용기 사전
- [**일괄 스타일 · TSV · 3D**](/ko/excel/tools-export) — 시트 전체 기존 수식 스타일 한 번에 통일하기
