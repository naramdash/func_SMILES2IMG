# 수식 오류 (`#NAME?`, `#VALUE!`, `#N/A`)

Microsoft Excel에서 `func_SMILES2IMG` 수식이 반환하는 오류·상태 코드별 원인과 해결 방법입니다.

## 1. `#NAME?` 오류

엑셀이 함수 이름을 인식하지 못할 때 나타납니다:

| 원인 | 확인 방법 | 해결 방법 |
| :--- | :--- | :--- |
| **Excel 2016 / 2019 / 2021** | 인셀 이미지 수식(`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`)은 Microsoft 365 및 Excel 2024 이상 전용입니다. | 뒤에 `.FLOAT`를 붙이세요: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT("Aspirin")` |
| **추가 기능 미로드 또는 차단** | `=SMILES2MW("CCO")`도 `#NAME?`이 뜨거나 `[SMILES]` 탭이 안 보입니다. | [설치 보안 차단](/ko/guide/troubleshooting/install) 및 [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)를 확인하세요. |

## 2. `#VALUE!` 오류

수식은 실행되었지만 입력값을 유효한 화학 구조로 파싱할 수 없을 때 나타납니다.

1. **화학 문법 확인:** 닫히지 않은 고리 번호(예: 끝 `1`이 빠진 `c1ccccc`), 괄호 불일치, 불필요한 줄바꿈·공백을 확인하세요.
2. **빈 셀 · 병합된 셀:** 병합된 셀을 참조할 때는 반드시 **가장 왼쪽 위(좌상단) 셀**을 지정하세요.
3. **정확한 파서 오류 확인 (`Diagnostics`):** 리본 **`[SMILES]` ➔ `[Diagnostics]`** 메뉴를 클릭하면 EPAM Indigo 파서가 보고한 정확한 문법 오류 메시지를 확인할 수 있습니다.

## 3. `#N/A`가 잠시(~0.1초) 표시됨

* **정상적인 비동기 표시:** 새 구조식을 처음 렌더링하거나 미캐시 이름을 PubChem에서 조회하는 동안 잠깐 `#N/A`가 떴다가 **완료 즉시 최종 이미지·값으로 자동 갱신**됩니다.
* **`=NAME2IMG` / `=NAME2SMILES`에서 `#N/A`가 유지되는 경우:** 내장 사전이나 PubChem에 없는 이름·오타거나 오프라인 상태일 수 있습니다. 영문 IUPAC명, CAS 번호(`"50-78-2"`), 또는 SMILES 문자열을 입력해 보세요.

## 관련 가이드

- [**이미지 크기 · .FLOAT**](/ko/excel/troubleshooting/image-sizing) — 작아 보이는 구조식 셀 크기 조절 및 Excel 2016~2021 `.FLOAT` 도형 관리
- [**40개 함수 전체 개요**](/ko/excel/formulas) — 정확한 수식 인자, 반환 데이터 형식 및 공용 스타일 키워드 확인
