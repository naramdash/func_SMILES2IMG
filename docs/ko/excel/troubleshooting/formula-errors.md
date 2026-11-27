# 수식 오류 (`#NAME?`, `#VALUE!`, `#N/A`)

Microsoft Excel에서 `func_SMILES2IMG` 수식이 반환하는 오류·상태 코드별 원인과 해결 방법입니다.

## 1. `#NAME?` 오류 및 구버전 Excel 안내

엑셀이 함수 이름을 인식하지 못할 때 `#NAME?`가 나타나며, 구버전 Excel 2016~2021에서는 `.FLOAT` 안내 메시지가 표시됩니다:

| 증상 | 원인 | 해결 방법 |
| :--- | :--- | :--- |
| **모든 수식에 `#NAME?` 표시** | 추가 기능이 로드되지 않았거나 보안 차단됨, 또는 `func_SMILES2IMG`가 없는 PC에서 재계산됨 (`=SMILES2MW("CCO")`도 `#NAME?`이 뜨거나 `[SMILES]` 탭이 안 보임). | [설치 보안 차단](/ko/guide/troubleshooting/install) 및 [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)를 확인하세요. |
| **`"Needs Excel 2024/365: use SMILES2IMG.FLOAT"`** | 인셀 이미지 수식(`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`)은 Microsoft 365 및 Excel 2024 이상 전용입니다. | 이미지 수식 뒤에 `.FLOAT`를 붙이세요: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT("Aspirin")` (물성·SAR 수식은 `.FLOAT` 없이 동작). |

## 2. `#VALUE!` 오류

수식은 실행되었지만 입력값을 유효한 화학 구조로 파싱할 수 없을 때 나타납니다.

1. **빈 셀 또는 글자 수 제한 초과:** 비어 있거나 공백만 있는 셀, 그리고 **16,384자**를 넘는 입력은 `#VALUE!`를 반환합니다.
2. **화학 문법 · 원자가 · 포맷 확인:** 닫히지 않은 고리 번호(예: 끝 `1`이 빠진 `c1ccccc`), 괄호 불일치·빈 괄호(`CC(=O`), 잘못된 원자가(예: 중성 5가 탄소 `C(C)(C)(C)(C)C`), `INCHI2*`에 일반 SMILES 전달, `KET2*`에 일반 문자열 전달, `*.MATCH`의 잘못된 SMARTS 패턴을 점검하세요.
3. **병합된 셀:** 병합된 셀을 참조할 때는 반드시 **가장 왼쪽 위(좌상단) 셀**을 지정하세요.
4. **정확한 파서 오류 확인 (`Diagnostics`):** 리본 **`[SMILES]` ➔ `[Diagnostics]`** 메뉴를 클릭하면 오류가 발생한 수식 함수명과 EPAM Indigo 파서가 보고한 정확한 문법 오류 메시지(예: `SMILES2IMG: Indigo: ...`)를 확인할 수 있습니다.

<FeatureMedia id="excel-diagnostics-dialog" />

## 3. `#N/A` 상태, 일시 빈 셀 표시 및 가드레일

* **비동기 렌더링 중 잠시 빈 셀 표시 (`...2IMG`):** 캐시되지 않은 구조식을 처음 렌더링하는 동안 `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG`는 대기열 작업이 끝날 때까지 잠시 빈 셀(`""`)로 표시된 뒤 **최종 이미지로 자동 갱신**됩니다.
* **알 수 없는 이름 · 오프라인 조회 · 잘못된 CAS 체크 디지트 (`#N/A`):** 내장 사전이나 PubChem에 없는 이름(또는 미캐시 오프라인 상태)이거나 수학적 체크 디지트가 틀린 CAS 번호(예: `"50-78-3"`)를 `=NAME2*` · `*2NAME`에 넣으면 `#N/A`를 반환합니다(체크 디지트가 틀린 CAS 번호는 네트워크 요청 없이 로컬에서 즉시 차단).
* **단일 분자 스칼라 함수의 쿼리 · 반응식 가드 (`#N/A`):** R-기(`*`, `[R1]`)나 반응식(`>>`)을 `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2FSP3` 같은 단일 분자 물성 함수에 전달하면 엉뚱한 숫자 대신 `#N/A`를 반환합니다.

## 관련 가이드

- [**이미지 크기 · .FLOAT**](/ko/excel/troubleshooting/image-sizing) — 작아 보이는 구조식 셀 크기 조절 및 Excel 2016~2021 `.FLOAT` 도형 관리
- [**49개 함수 전체 개요**](/ko/excel/formulas) — 정확한 수식 인자, 반환 데이터 형식 및 공용 스타일 키워드 확인
