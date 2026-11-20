# 구버전 Excel (`.FLOAT`)

네이티브 인셀 이미지 수식(`=SMILES2IMG`, `=NAME2IMG`, `=KET2IMG`, `=INCHI2IMG`)은 **Microsoft 365 및 Excel 2024 이상** 전용입니다. **Excel 2016, 2019, 2021**에서는 뒤에 `.FLOAT`를 붙여 셀에 고정되는 부동 그림 도형으로 렌더링합니다.

## `.FLOAT` 수식 대응표

| 최신 Excel (365 / 2024+) | 구버전 Excel (2016 / 2019 / 2021) | 사용 예시 |
| :--- | :--- | :--- |
| `=SMILES2IMG(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2, [bg], [style], [transform])` | `=SMILES2IMG.FLOAT(A2)` |
| `=NAME2IMG(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT(name, [bg], [style], [transform])` | `=NAME2IMG.FLOAT("Caffeine")` |
| `=KET2IMG(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2, [bg], [style], [transform])` | `=KET2IMG.FLOAT(A2)` |
| `=INCHI2IMG(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2, [bg], [style], [transform])` | `=INCHI2IMG.FLOAT(A2)` |

<FeatureMedia id="excel-legacy-float" />

## `.FLOAT` 도형 동작 특징

* 생성된 그림 도형은 대상 셀에 고정(`xlMoveAndSize`)되어 행 높이·열 너비를 조절하면 비율을 유지하며 함께 확대·축소됩니다.
* `"bw"`, `"white"`, `"hl:..."`, `"num stereo"`, `"ink=#HEX"` 등 모든 스타일 옵션이 인셀 수식과 동일하게 작동합니다.
* 물성·Lipinski·SAR·변환 함수(`=SMILES2MW`, `=SMILES2LIPINSKI`, `=SMILES.MATCH` 등)는 일반 숫자/텍스트를 반환하므로 **2016~365 모든 Excel 버전**에서 `.FLOAT` 없이 그대로 사용합니다.
