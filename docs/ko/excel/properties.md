# 물성 · Lipinski 약물성

분자량, HRMS 정밀질량, LogP, TPSA, Lipinski Rule of 5 통과 여부, 3D 입체 복잡도(`Fsp3`, 키랄 중심)를 엑셀 시트에서 0ms 지연 없이 로컬 계산합니다.

```text
=SMILES2MW(A2)
=SMILES2EXACTMASS(A2)
=SMILES2LIPINSKI(A2)
=SMILES2FSP3(A2)
```

<FeatureMedia id="excel-cheminformatics" />

## 물리화학적 물성 · 약물성 수식 목록

| 물성 항목 | 엑셀 수식 | 의약화학 기준 | 아스피린 (`A2`) | 비고 |
| :--- | :--- | :---: | :---: | :--- |
| **화학식 (Hill)** | `=SMILES2FORMULA(A2)` | — | `C9H8O4` | 표준 힐 표기법 |
| **분자량 (MW)** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | 당량 곱셈 가능한 숫자(`double`) |
| **정밀질량 (HRMS)** | `=SMILES2EXACTMASS(A2)` | — | `180.0423` | `[M+H]+` = `+1.0073`, `[M+Na]+` = `+22.9892` |
| **LogP / TPSA** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | **≤ 5.0** / **≤ 140 Å²** | `1.31` / `63.60` | CNS BBB 기준 ≤ 90 Å² |
| **HBD / HBA / RotB** | `=SMILES2HBD` / `HBA` / `ROTBONDS` | **≤ 5** / **≤ 10** / **≤ 10** | `1` / `4` / `3` | 경구 흡수성 지표 |
| **Lipinski 판정** | `=SMILES2LIPINSKI(A2)` | — | `PASS` | 4대 조건 종합 판정 |
| **중원자 / 고리 / 방향족** | `=SMILES2HEAVY` / `RINGS` / `AROMATIC` | **방향족 ≤ 3** | `13` / `1` / `1` | 용해도·개발성 추적 |
| **3D Fsp³ / 키랄 중심** | `=SMILES2FSP3(A2)` / `=SMILES2CHIRAL(A2)` | **Fsp³ ≥ 0.25** | `0.11` / `0` | 3차원 입체 복잡도 지표 |

::: tip InChI나 KET 입력을 쓰는 경우
`SMILES2` 대신 `INCHI2` 또는 `KET2`로 바꾸면 동일하게 동작합니다(예: `=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).
:::

## Excel `IF` / `AND` 스크리닝 결합 예시

| 스크리닝 조건 | 복사해서 쓰는 수식 (`A2` = SMILES) |
| :--- | :--- |
| **CNS BBB 후보 (`TPSA ≤ 90`, `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "CNS 후보", "일반")` |
| **3D 입체 리드 (`Fsp3 ≥ 0.25`, `방향족 ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "입체 우수", "평면")` |
