# 표기법 변환 · InChIKey 중복 제거

**물질명**, **CAS 등록번호**, **표준 SMILES**, **IUPAC InChI**, **27자리 고정 길이 `InChIKey` 해시** 간 상호 변환 함수를 제공해, 시약 업체나 문헌마다 달라지는 분자 문자열을 표준화하고 `XLOOKUP` 병합 및 중복 화합물 제거를 자동화합니다.

<FormulaTable
  title="화학 식별자 상호 변환 시트"
  :columns='["입력값 (SMILES / InChI)", "수식"]'
  :rows='[
    { input: "", formula: "=NAME2SMILES(\"50-78-2\")", desc: "CAS 번호·시약명 단독 입력으로 SMILES 변환", out: "CC(=O)Oc1ccccc1C(=O)O" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2INCHI(A3)", desc: "SMILES·KET 셀을 IUPAC InChI로 변환", out: "InChI=1S/..." },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2INCHIKEY(A4)", desc: "중복 제거용 27자리 고정 InChIKey 생성", out: "BSYNRYMUTXBXSQ-..." },
    { input: "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3", formula: "=INCHI2SMILES(A5)", desc: "IUPAC InChI를 표준 SMILES로 역변환", out: "CCO" }
  ]'
/>

<FeatureMedia id="excel-conversions" />

## 1. 화학 표기법 상호 변환 함수 목록

모든 변환 함수는 엑셀 텍스트(`String`)를 반환합니다(`*2NAME` / `NAME2*` 는 내장 사전 + 로컬 캐시 PubChem 조회, SMILES · InChI · KET 상호 변환은 0ms 순수 오프라인):

| 변환 방향 | 엑셀 수식 | 결과 예시 (`Aspirin`) |
| :--- | :--- | :--- |
| **시약 이름 / CAS 번호 ➔ 표준 SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ 물질명** | `=SMILES2NAME(A2)` · `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `Aspirin` |
| **SMILES / KET ➔ IUPAC InChI** | `=SMILES2INCHI(A2)` · `=KET2INCHI(A2)` | `InChI=1S/C9H8O4/c1-6(10)...` |
| **InChI ➔ 표준 SMILES** | `=INCHI2SMILES(A2)` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` · `=INCHI2INCHIKEY(A2)` · `=KET2INCHIKEY(A2)` | `BSYNRYMUTXBXSQ-UHFFFAOYSA-N` |

## 2. `InChIKey` 로 서로 다른 DB 간 중복 화합물 100% 잡아내기

같은 분자라도 시약 공급처나 문헌에 따라 SMILES 원자 나열 순서가 다를 수 있어(`c1ccccc1C(=O)O` vs `O=C(O)c1ccccc1`) 일반 텍스트 비교로는 중복을 놓치기 쉽습니다.

보조 열에 원자 순서와 무관한 27자리 고정 해시 `=SMILES2INCHIKEY(B2)` 를 생성해 두면 엑셀 표준 도구만으로 중복 제거와 시트 병합이 완벽해집니다:

| 실무 작업 | 바로 쓰는 엑셀 레시피 (`D`열 = `=SMILES2INCHIKEY(B2)`) | 동작 결과 |
| :--- | :--- | :--- |
| **중복 화합물 실시간 표시** | `=IF(COUNTIF($D$2:$D$500, D2)>1, "중복", "고유")` | 서로 다른 표기의 동일 화합물을 즉시 표시 |
| **다른 공급처 시트 정확한 조회** | `=XLOOKUP(D2, VendorB!$D:$D, VendorB!$E:$E, "없음")` | SMILES 문자열이 달라도 같은 분자 기준으로 활성·단가 매칭 |
| **마스터 라이브러리 중복 제거** | 표 선택 ➔ Excel **`[데이터]` ➔ `[중복된 항목 제거]`** 에서 `InChIKey` 열 선택 | 화학적으로 동일한 중복 행을 오차 없이 정리 |

## 관련 가이드

- [**부분구조 · 유사도 검색**](/ko/excel/sar) — `=SMILES.MATCH` 골격 필터링 및 Tanimoto 유사도 순위 산출
- [**SDF 파일 입출력**](/ko/excel/sdf) — 외부 `.sdf`/`.mol` 파일을 12열 라이브 시트로 불러오거나 시트를 `.sdf` 로 내보내기
- [**40개 함수 전체 개요**](/ko/excel/formulas) — 전체 엑셀 수식 규격 및 반환 타입 한눈에 보기
