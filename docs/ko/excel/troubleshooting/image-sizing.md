# 이미지 크기 · `.FLOAT` 모드

엑셀 셀 안 구조식 크기 조절, Excel 2016~2021 `.FLOAT` 도형 동작, `.xlsx` 파일 외부 공유 시 주의사항입니다.

## 1. 셀 안의 분자 그림이 너무 작게 보임

인셀 이미지는 화학 구조 비율이 찌그러지지 않도록 가로세로 비율을 고정한 채 셀 크기에 맞춰 자동 축소·확대됩니다.

* **해결:** 해당 셀의 **행 높이**(권장 `62~100 pt`)와 **열 너비**를 넓혀 주세요.
* **일괄 팁:** 화합물 행 전체를 한꺼번에 드래그 선택한 뒤 행 경계선 하나만 조절하면 수백 행의 구조식이 동시에 선명하게 커집니다.

## 2. Excel 2016 / 2019 / 2021 (`.FLOAT` 도형)

구버전 영구 라이선스 엑셀은 네이티브 인셀 이미지 수식(`=SMILES2IMG`)을 지원하지 않습니다.

* 대신 `.FLOAT` 수식을 사용하세요: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT("Aspirin")`, `=KET2IMG.FLOAT(A2)`, `=INCHI2IMG.FLOAT(A2)`
* 생성된 그림 도형은 셀에 고정(`xlMoveAndSize`)되어 셀 크기를 조정하면 구조식도 함께 조절됩니다.

## 3. `#NAME?` 없이 `.xlsx` 파일 공유하기

* 캐시된 셀 이미지와 계산된 물성값은 재계산(<kbd>F9</kbd>)이 일어나기 전까지 다른 PC에서도 그대로 표시됩니다.
* `func_SMILES2IMG`가 없는 동료에게 최종 배포할 때는 숫자·텍스트 수식 열을 복사(<kbd>Ctrl</kbd> + <kbd>C</kbd>) ➔ **값으로 붙여넣기**하거나 **PDF로 저장**해 전달하세요. 자세한 내용은 [문서 공유](/ko/guide/troubleshooting/document-sharing)를 참고하세요.

## 관련 가이드

- [**엑셀 수식 오류**](/ko/excel/troubleshooting/formula-errors) — #NAME?, #VALUE!, #N/A 및 Diagnostics 파서 오류 해결
- [**Excel 레시피 1~5**](/ko/excel/) — 구조식 표, 이름/CAS 조회, 스타일, .FLOAT 모드 실무 예제
