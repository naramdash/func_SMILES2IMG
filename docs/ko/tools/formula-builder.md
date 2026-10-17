# 엑셀 함수 표현 생성기 & 스타일 시뮬레이터

복잡한 수식 문법이나 옵션 키워드를 외울 필요 없이, 원하는 옵션을 마우스로 클릭해 **실시간으로 구조식 렌더링 결과를 미리 보고 완성된 엑셀 수식을 원클릭으로 복사**할 수 있는 웹 도구 모음입니다.

---

## 1. 🧪 인터랙티브 엑셀 수식 표현 생성기

내 입력 데이터 종류(`SMILES`, `시약 이름`, `CAS 번호`, `반응식 KET`, `InChI`), 엑셀 버전, 배경색, 선 색상, 부분구조 하이라이트(`hl:`), 원자 번호(`num`) 옵션을 클릭해 조합해 보세요:

<FormulaBuilder />

---

## 2. 🎨 일괄 스타일(`Batch Style`) 프리셋 시뮬레이터

Excel, Word, PowerPoint 상단 리본의 **`[SMILES]`** ➔ **`[Batch Style]`** 메뉴에서 제공하는 4가지 문서 전체 스타일 프리셋을 클릭해 비교해 보세요:

<BatchStylePreview />

---

## 3. ⚡ 자주 쓰는 물성 · 약물성 · SAR 수식 빠른 복사

구조식 이미지 함수와 함께 옆 열에 바로 붙여넣어 쓸 수 있는 핵심 계산 수식 모음입니다 (`A2` 셀 기준):

| 구하고 싶은 값 | 복사할 엑셀 수식 | 결과 예시 (아스피린) |
| :--- | :--- | :--- |
| **화학식 (Hill Formula)** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **평균 분자량 (MW, g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **정밀 질량 (Exact Mass, Da)** | `=SMILES2EXACTMASS(A2)` | `180.0423` |
| **Lipinski Rule of 5 판정** | `=SMILES2LIPINSKI(A2)` | `PASS` |
| **지용성 (LogP) / 극성 표면적 (TPSA)** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | `1.31` / `63.60` |
| **특정 작용기 포함 여부 (예: 카복실산)** | `=SMILES.MATCH(A2, "[CX3](=O)[OX2H1]")` | `TRUE` |
| **기준 물질(`$D$1`)과의 구조 유사도** | `=SMILES.SIMILARITY(A2, $D$1)` | `0.85` (0.0 ~ 1.0) |

---

## 4. 🔍 12대 필수 작용기(SMARTS) 검색 & 하이라이트 수식 빠른 복사

원하는 작용기의 **포함 여부 필터(`=SMILES.MATCH`)** 수식이나 **하이라이트 이미지(`=SMILES2IMG`)** 수식을 그대로 복사해 시트에 붙여넣으세요:

| 찾을 작용기 / 고리 골격 | 포함 여부 판정 (`TRUE` / `FALSE`) | 해당 작용기 컬러 하이라이트 구조식 |
| :--- | :--- | :--- |
| **벤젠 고리 (방향족)** | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **피리딘 고리 (헤테로방향족)** | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **카복실산 (`-COOH`)** | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **에스터 (`-COO-`)** | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **아마이드 (`-CONH-`)** | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **설폰아마이드 (`-SO₂NH-`)** | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **케톤 (`-C(=O)-`)** | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **알데하이드 (`-CHO`)** | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **페놀 (방향족 `-OH`)** | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **나이트로기 (`-NO₂`)** | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **할로겐 원소 (`F, Cl, Br, I`)** | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **트라이플루오로메틸 (`-CF₃`)** | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

---

## 🔗 관련 가이드 바로가기

<div class="feature-grid-2">
  <div class="feature-card">
    <h4>📐 <a href="/ko/reference/formula">전체 엑셀 함수 치트시트</a></h4>
    <p>25종 이상의 모든 엑셀 화학 함수와 인수 규격, 스타일 키워드 전체 목록을 확인하세요.</p>
  </div>
  <div class="feature-card">
    <h4>📗 <a href="/ko/guide/excel">Excel 기초 가이드</a></h4>
    <p>엑셀 시트에서 수식으로 대량 구조식 표를 만들고 더블클릭으로 편집하는 방법을 알아보세요.</p>
  </div>
</div>
