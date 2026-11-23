# SMILES · InChI 구조식 렌더링

**SMILES 문자열(`=SMILES2IMG`)** 또는 **IUPAC InChI 문자열(`=INCHI2IMG`)** 열을 엑셀 셀 안에서 선명한 2D 화학 구조식 표로 자동 변환합니다. 행 높이에 맞춰 비율이 깨지지 않고 확대되며 표 정렬·필터와도 그대로 연동됩니다.

<FormulaTable
  title="SMILES · InChI 구조식 렌더링 시트"
  :columns='["화합물명", "SMILES / InChI", "수식"]'
  :rows='[
    { cells: ["Ethanol", "CCO", "=SMILES2IMG(B2)"], desc: "B열 SMILES 문자열로 인셀 2D 구조식 렌더링", out: "에탄올 2D" },
    { cells: ["Ethanol", "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3", "=INCHI2IMG(B3)"], desc: "B열 IUPAC InChI 문자열로 인셀 2D 구조식 렌더링", out: "에탄올 2D" }
  ]'
/>

<FeatureMedia id="excel-formula-table" />

## 1. SMILES · InChI 공통 4인수 문법

두 수식 모두 동일한 **4개 인수 규격 `(입력, [배경색], [스타일], [변형])`** 을 사용하며 오프라인으로 동작합니다:

| 입력 데이터 | Microsoft 365 / Excel 2024+ | Excel 2016~2021용 | 특징 |
| :--- | :--- | :--- | :--- |
| **SMILES 문자열** | `=SMILES2IMG(B2, [배경색], [스타일], [변형])` | `=SMILES2IMG.FLOAT(...)` | KET JSON 입력도 자동 판별 |
| **IUPAC InChI** (`InChI=1S/...`) | `=INCHI2IMG(C2, [배경색], [스타일], [변형])` | `=INCHI2IMG.FLOAT(...)` | PubChem·ChEMBL·NIST 표준 InChI 직접 렌더링 |

## 2. 수백 행 구조식 표 만들기

1. `A`열에 화합물 이름, `B`열에 SMILES 문자열(예: 아스피린 `CC(=O)Oc1ccccc1C(=O)O`)을 입력합니다.
2. `C2` 셀에 `=SMILES2IMG(B2)` (또는 `=INCHI2IMG(B2)`)를 입력하고 <kbd>Enter</kbd> 를 누릅니다 (Excel 2016~2021은 뒤에 **`.FLOAT`** 추가).
3. **행 높이**(`62~80 pt` 권장)와 열 너비를 넓히면 결합 비율이 유지된 채 확대되며, 표 정렬·필터 시 이미지도 행과 함께 정확히 이동합니다.
4. 채우기 핸들을 아래로 드래그해 전체 열을 일괄 렌더링합니다.

## 3. 직접 그리기 · 붙여넣기 변환 · 더블클릭 수정

| 작업 | 조작 방법 | 동작 결과 |
| :--- | :--- | :--- |
| **새 분자 그리기** | 셀 선택 ➔ **`[SMILES]` ➔ `[Insert / Edit (SMILES)]`** (또는 `InChI`) ➔ Ketcher 작성 후 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> | 활성 셀에 `=SMILES2IMG("...")` (또는 `=INCHI2IMG(...)`) 자동 입력 |
| **텍스트/클립보드 변환** | SMILES/InChI/KET/이름/CAS 복사 또는 텍스트 셀 선택 ➔ **`[SMILES]` ➔ `[Paste Structure]`** 클릭 | `=SMILES2IMG(...)` 또는 `=KET2IMG(...)` 수식으로 변환 (일반 물질명·모호한 짧은 영단어는 확인 후 삽입) |
| **더블클릭 재편집** | 렌더링된 구조식 셀 또는 `.FLOAT` 그림 **더블클릭** | 기존 분자가 로드된 채 **Ketcher 2D** 실행 (모호한 일반 단어 셀에서는 빈 캔버스로 열림) |

## 관련 가이드

- [**시약명 · 용매 · CAS**](/ko/excel/name-cas) — `=NAME2IMG` 로 영문 이름, `THF`/`DMSO`, `"50-78-2"` 그리기 (내장 오프라인 사전 + PubChem 30일 로컬 캐시)
- [**합성 반응식 (KET)**](/ko/excel/reactions) — `=KET2IMG` 로 화살표·촉매 포함 반응식 무손실 저장
- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — `[배경색]`, `[스타일]` (`"bw"`, `"hl:..."`, `"num stereo"`), `[변형]` 옵션 활용
- [**구버전 Excel (.FLOAT)**](/ko/excel/legacy-float) — Excel 2016~2021용 셀 고정 도형 수식 안내
