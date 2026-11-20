# 표기법 변환 (InChI · CAS)

화합물 이름, CAS 번호, 표준 SMILES, IUPAC InChI, 27자리 InChIKey 간 상호 변환으로 데이터 정규화 및 중복 제거(`XLOOKUP` / `VLOOKUP`)를 자동화합니다.

```text
=NAME2SMILES("50-78-2")
=SMILES2INCHI(B2)
=SMILES2INCHIKEY(B2)
=INCHI2IMG(C2)
```

<FeatureMedia id="excel-conversions" />

## 표기법 상호 변환 수식

| 변환 방향 | 엑셀 수식 | 결과 예시 |
| :--- | :--- | :--- |
| **이름 / CAS 번호 ➔ SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES ➔ 물질명** | `=SMILES2NAME(A2)` | `Aspirin` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` / `=KET2INCHI(A2)` | `InChI=1S/C9H8O4/c1-6(10)...` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `CC(=O)Oc1ccccc1C(=O)O` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` / `=INCHI2INCHIKEY(A2)` / `=KET2INCHIKEY(A2)` | `BSYNRYMUTXBXSQ-UHFFFAOYSA-N` |

::: tip 서로 다른 DB의 중복 화합물 제거 팁
DB마다 같은 분자의 SMILES 문자열 표기가 다를 수 있으므로 보조 열에 `=SMILES2INCHIKEY(A2)`를 생성한 뒤 엑셀의 `XLOOKUP`이나 **중복된 항목 제거**를 적용하면 정확하게 일치시킬 수 있습니다.
:::
