# 엑셀 함수 치트시트 (Formula Cheat Sheet)

엑셀 시트에서 바로 복사해 쓸 수 있는 **25종 이상의 화학 구조식 렌더링, 분자량·약물성 계산, 부분구조 검색 수식**을 목적별로 정리한 빠른 사전입니다.

---

## 1. 🖼️ 구조식 이미지 그리기 (`...2IMG`)

셀 안에 2D 화학 구조식이나 합성 반응식 그림을 그려 넣는 함수입니다.

| 내가 가진 입력 데이터 | 사용할 수식 (Microsoft 365 / Excel 2024+) | 구버전 엑셀용 (Excel 2016 ~ 2021) | 수식 예시 |
| :--- | :--- | :--- | :--- |
| **SMILES 문자열** | `=SMILES2IMG(셀, [배경색], [옵션])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **시약 이름 또는 CAS 번호** | `=NAME2IMG(이름_또는_CAS, [배경색], [옵션])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **화학 반응식 (KET)** | `=KET2IMG(셀, [배경색], [옵션])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |
| **InChI 문자열** | `=INCHI2IMG(셀, [배경색], [옵션])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |

::: tip 💡 배경색(`[배경색]`)과 옵션(`[옵션]`)은 생략해도 됩니다!
* 기본값으로 쓰려면 `=SMILES2IMG(A2)`처럼 셀 주소 하나만 넣으면 됩니다.
* 옵션을 주고 싶을 때는 두 번째 인수(배경색)를 빈 문자열(`""`)로 두고 세 번째 인수에 옵션을 적습니다:  
  예: `=SMILES2IMG(A2, "", "bw")` *(논문용 흑백 선화)*
:::

---

## 2. 🏷️ 이름 · 화학식 · 분자량 구하기

모든 계산(물질명 웹 조회 제외)은 인터넷 연결 없이 **0ms 순수 오프라인**으로 즉시 계산됩니다.

| 구하고 싶은 값 | 수식 | 결과 예시 (`A2` = 아스피린) | 설명 |
| :--- | :--- | :--- | :--- |
| **물질 일반명 (Name)** | `=SMILES2NAME(A2)` | `"Aspirin"` | IUPAC / PubChem 기반 영문 물질명 조회 |
| **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | `"C9H8O4"` | 표준 Hill 표기법 화학식 |
| **분자량 (MW)** | `=SMILES2MW(A2)` | `180.16` | 표준 분자량 (g/mol) |
| **정밀 질량 (Exact Mass)** | `=SMILES2EXACTMASS(A2)` | `180.0423` | 질량분석(HRMS/MS)용 단일동위원소 질량 (Da) |

> 🔄 **입력이 InChI나 KET 반응식이라면?**  
> 함수 앞글자의 `SMILES2`만 `INCHI2` 또는 `KET2`로 바꾸면 똑같이 작동합니다!  
> 예: `=INCHI2MW(A2)`, `=INCHI2FORMULA(A2)`, `=KET2MW(A2)`

---

## 3. 💊 약물성(Lipinski Rule of 5) & 물리화학 지표

신약 후보 물질의 경구 흡수 가능성과 물리화학적 성질을 평가하는 수식입니다.

| 구하고 싶은 값 | 수식 | 결과 예시 | 기준 및 설명 |
| :--- | :--- | :--- | :--- |
| **Lipinski 약물성 판정** | `=SMILES2LIPINSKI(A2)` | `"PASS"` / `"FAIL"` | Lipinski's Rule of 5 통과 여부 자동 판정 |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | `1.31` | Wildman-Crippen 옥탄올-물 분배계수 |
| **극성 표면적 (TPSA)** | `=SMILES2TPSA(A2)` | `63.60` | 위상 극성 표면적 (Å²) |
| **수소 결합 공여체 수 (HBD)** | `=SMILES2HBD(A2)` | `1` | $-OH$, $-NH$ 등 수소 결합 공여체 개수 |
| **수소 결합 수용체 수 (HBA)** | `=SMILES2HBA(A2)` | `4` | 질소($N$), 산소($O$) 등 수소 결합 수용체 개수 |
| **회전 가능한 결합 수** | `=SMILES2ROTBONDS(A2)` | `3` | 단일 회전 가능 결합(Rotatable bonds) 개수 |

---

## 4. 🔍 부분구조 검색 & 유사도 비교 (SAR 스크리닝)

수천 개의 화합물 시트에서 특정 골격을 찾거나 리드 화합물과의 유사도를 순위 매길 때 사용합니다.

| 목적 | 수식 | 결과 예시 | 활용 설명 |
| :--- | :--- | :--- | :--- |
| **특정 부분구조 포함 여부** | `=SMILES.MATCH(대상셀, "찾을구조")` | `TRUE` / `FALSE` | 대상 분자 안에 지정한 SMILES/SMARTS 골격이 들어있으면 `TRUE` 반환 |
| **두 분자 간 2D 유사도 점수** | `=SMILES.SIMILARITY(분자1, 분자2, [방식])` | `0.0` ~ `1.0` | 두 분자의 2D 지문 유사도 계산 (`"tanimoto"` 기본값, `"tversky"` 선택 가능) |

---

## 5. 🔄 화학 표기법 상호 변환 (이름 · CAS · SMILES · InChI)

| 변환 방향 | 수식 | 결과 예시 |
| :--- | :--- | :--- |
| **시약 이름 / CAS 번호 ➔ SMILES** | `=NAME2SMILES("50-78-2")` 또는 `=NAME2SMILES("Aspirin")` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES ➔ InChI** | `=SMILES2INCHI(A2)` | `"InChI=1S/C9H8O4/c1-6(10)..."` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES ➔ InChIKey (27자리 해시)** | `=SMILES2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

---

## 6. 🎨 이미지 스타일 옵션 키워드 (`options`)

`=SMILES2IMG(A2, [배경색], [옵션])`의 세 번째 인수에 아래 키워드들을 공백이나 쉼표로 자유롭게 조합해 넣을 수 있습니다 (예: `"bw num stereo"`).

### 색상 및 강조 옵션
| 옵션 키워드 | 효과 | 수식 예시 |
| :--- | :--- | :--- |
| `colored` | 원소별 표준 컬러 표시 (기본값) | `=SMILES2IMG(A2, "", "colored")` |
| `bw` | 학술지(ACS) 및 특허용 고대비 흑백 선화 | `=SMILES2IMG(A2, "", "bw")` |
| `hl:<구조>` | 지정한 SMILES/SMARTS 부분구조를 컬러로 하이라이트 | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| `white` | 어두운 배경 위에서 선명하게 보이는 흰색 결합선 | `=SMILES2IMG(A2, "#1e1e1e", "white")` |
| `ink=#HEX` | 원하는 색상 코드(HEX)로 결합선 색상 지정 | `=SMILES2IMG(A2, "", "ink=#0055FF")` |

### 원자 라벨 및 회전/반전 옵션
| 옵션 키워드 | 효과 | 수식 예시 |
| :--- | :--- | :--- |
| `num` | NMR/MS 피크 지정을 위한 원자 번호(1, 2, 3...) 표시 | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | 카이랄 중심 입체화학 라벨(*R/S*, *E/Z*) 표시 | `=SMILES2IMG(A2, "", "stereo")` |
| `h` | 생략된 수소 원자($H$) 모두 펼쳐서 표시 | `=SMILES2IMG(A2, "", "h")` |
| `all-c` | 모든 탄소 골격 꼭짓점에 `C` 원소 기호 표시 | `=SMILES2IMG(A2, "", "all-c")` |
| `flip` / `flipv` | 분자를 좌우(`flip`) 또는 상하(`flipv`)로 대칭 반전 | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` / `rot270` | 분자를 시계 방향으로 90°, 180°, 270° 회전 | `=SMILES2IMG(A2, "", "rot90")` |

---

## 7. 📋 복사해서 바로 쓰는 실전 수식 모음

```text
# 1. CAS 번호나 이름으로 바로 구조식 그리기
=NAME2IMG("50-78-2")
=NAME2IMG("Caffeine")

# 2. 벤젠 고리(c1ccccc1) 부분만 하이라이트해서 그리기
=SMILES2IMG(A2, "", "hl:c1ccccc1")

# 3. 특허/논문 투고용 흑백 선화 + 입체화학(R/S) 표시
=SMILES2IMG(A2, "", "bw stereo")

# 4. 다크 모드 시트용 (어두운 배경 + 흰색 선)
=SMILES2IMG(A2, "#1e1e1e", "white")

# 5. 특정 작용기(카복실산 C(=O)O) 포함 여부 확인 (TRUE / FALSE)
=SMILES.MATCH(A2, "C(=O)O")

# 6. B1 셀의 기준 화합물과 얼마나 비슷한지 유사도(0~1) 계산
=SMILES.SIMILARITY(A2, $B$1)
```
