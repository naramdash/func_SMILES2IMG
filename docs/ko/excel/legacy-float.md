# 구버전 Excel (`.FLOAT`)

네이티브 인셀 이미지 수식(`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`)은 **Microsoft 365 및 Excel 2024 이상** 전용입니다. **Excel 2016, 2019, 2021**에서는 뒤에 **`.FLOAT`** 를 붙여 셀에 고정되는 부동 그림 도형으로 렌더링하세요.

<FormulaTable
  title="Excel 2016~2021 (.FLOAT) 시트"
  :columns='["SMILES / 입력값", "수식"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG.FLOAT(A2)", desc: "SMILES 셀로 셀 고정 부동 그림 생성", out: "xlMoveAndSize" },
    { input: "", formula: "=NAME2IMG.FLOAT(\"Caffeine\")", desc: "화합물명 단독 부동 구조식 도형 생성", out: "xlMoveAndSize" }
  ]'
/>

<FeatureMedia id="excel-legacy-float" />

## 1. `.FLOAT` 수식 대응표

| 최신 Excel (365 / 2024+) | 구버전 Excel (2016 / 2019 / 2021) | 사용 예시 |
| :--- | :--- | :--- |
| `=SMILES2IMG(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2)` |
| `=NAME2IMG(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT("Caffeine")` |
| `=KET2IMG(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2)` |
| `=INCHI2IMG(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2)` |

## 2. `.FLOAT` 도형 동작 특징

* 생성된 그림 도형은 대상 셀에 고정(`xlMoveAndSize`)되어 행 높이·열 너비를 조절하면 비율을 유지하며 함께 확대·축소됩니다.
* `"bw"`, `"white"`, `"hl:..."`, `"num stereo"`, `"ink=#HEX"` 등 공용 스타일 키워드가 인셀 수식과 100% 똑같이 적용됩니다.
* 물성·Lipinski·SAR·표기법 변환 수식(`=SMILES2MW`, `=SMILES2LIPINSKI`, `=SMILES.MATCH`, `=NAME2SMILES` 등)은 일반 숫자/텍스트를 반환하므로 **2016~365 모든 Excel 버전**에서 `.FLOAT` 없이 그대로 작동합니다.

## 관련 가이드

- [**SMILES · InChI**](/ko/excel/structures) — 대량 구조식 표 만들기 및 더블클릭 Ketcher 2D 재편집
- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — 모든 `...2IMG` 및 `.FLOAT` 공용 스타일 옵션 매트릭스
- [**이미지 크기 · .FLOAT**](/ko/excel/troubleshooting/image-sizing) — 셀 크기 조절 및 `.xlsx` 파일 외부 공유 시 주의사항
