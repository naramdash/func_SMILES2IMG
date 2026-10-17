# 인터랙티브 엑셀 수식 빌더 & 함수 치트시트

원하는 옵션을 마우스로 클릭해 **실시간으로 엑셀 수식을 만들고 복사**하거나, 아래의 **목적별 치트시트 표**에서 25종 이상의 모든 함수를 빠르게 찾아보세요.

<FormulaBuilder />

---

## 1. 🖼️ 구조식 이미지 그리기 (`...2IMG`)

셀 안에 2D 화학 구조식이나 다단계 합성 반응식 그림을 그려 넣는 함수입니다.

| 내가 가진 입력 데이터 | 사용할 수식 (Microsoft 365 / Excel 2024+) | 구버전 엑셀용 (Excel 2016 ~ 2021) | 수식 예시 |
| :--- | :--- | :--- | :--- |
| **SMILES 문자열** | `=SMILES2IMG(셀, [배경색], [옵션], [변형])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **시약 이름 또는 CAS 번호** | `=NAME2IMG(이름_또는_CAS, [배경색], [옵션], [변형])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **화학 반응식 (KET)** | `=KET2IMG(셀, [배경색], [옵션], [변형])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |
| **InChI 문자열** | `=INCHI2IMG(셀, [배경색], [옵션], [변형])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |

::: tip 💡 빠른 단축 문법 팁
* **기본 컬러 + 투명 배경:** `=SMILES2IMG(A2)`처럼 첫 번째 인수만 적으면 됩니다.
* **초간단 흑백 변환:** 두 번째 인수 자리에 바로 `"bw"`(또는 `FALSE`)를 적으면 세 번째 인수를 생략하고도 즉시 흑백 선화가 됩니다 (`=SMILES2IMG(A2, "bw")`).
* **자유로운 구분자:** 세 번째 `[옵션]` 인수 안에서 여러 키워드를 조합할 때 공백(` `), 쉼표(`,`), 파이프(`|`), 세미콜론(`;`), 플러스(`+`) 중 어떤 기호를 써도 모두 인식합니다 (`"bw num stereo"` 또는 `"bw, num, stereo"`).
:::

---

## 2. 🏷️ 이름 · 화학식 · 분자량 · 정밀질량 구하기

물질명 온라인 조회(`=SMILES2NAME`)를 제외한 모든 계산은 인터넷 연결 없이 **0ms 순수 오프라인**으로 즉시 계산됩니다.

| 구하고 싶은 값 | SMILES 기준 수식 | InChI / KET 입력 시 | 결과 예시 (`A2` = 아스피린) | 설명 |
| :--- | :--- | :--- | :--- | :--- |
| **물질 일반명 (Name)** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)`<br>`=KET2NAME(A2)` | `"Aspirin"` | IUPAC / PubChem 기반 영문 물질명 조회 (로컬 캐시 지원) |
| **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)`<br>`=KET2FORMULA(A2)` | `"C9H8O4"` | 표준 Hill 표기법 화학식 |
| **분자량 (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)`<br>`=KET2MW(A2)` | `180.16` | 시약 당량 계산용 평균 분자량 (g/mol, 소수점 2자리) |
| **정밀 질량 (Exact Mass)** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)`<br>`=KET2EXACTMASS(A2)` | `180.0423` | 질량분석(HRMS/LC-MS) 피크 매칭용 단일동위원소 질량 (Da, 소수점 4자리) |

---

## 3. 💊 약물성(Lipinski Rule of 5) & 물리화학(ADME) 지표

신약 후보 물질의 경구 흡수 가능성과 물리화학적 성질을 평가하는 수식입니다. 입력 셀에 SMILES, InChI, KET 중 무엇이 들어있든 자동으로 인식합니다.

| 구하고 싶은 값 | 수식 | 결과 예시 | Lipinski 기준 및 설명 |
| :--- | :--- | :--- | :--- |
| **Lipinski 약물성 판정** | `=SMILES2LIPINSKI(A2)` | `"PASS"` / `"FAIL"` | 아래 4대 기준(MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10) 통과 여부 |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | `1.31` | Wildman-Crippen 옥탄올-물 분배계수 (≤ 5.0 권장) |
| **극성 표면적 (TPSA)** | `=SMILES2TPSA(A2)` | `63.60` | 위상 극성 표면적 (Å², 장내 흡수 ≤ 140, 뇌혈관장벽 투과 ≤ 90) |
| **수소 결합 공여체 수 (HBD)** | `=SMILES2HBD(A2)` | `1` | `-OH`, `-NH` 등 수소 결합 공여체 개수 (≤ 5 권장) |
| **수소 결합 수용체 수 (HBA)** | `=SMILES2HBA(A2)` | `4` | 질소(`N`), 산소(`O`) 등 수소 결합 수용체 개수 (≤ 10 권장) |
| **회전 가능한 결합 수** | `=SMILES2ROTBONDS(A2)` | `3` | 단일 회전 가능 결합 개수 (≤ 10 권장) |

---

## 4. 🔍 부분구조 검색 & 유사도 비교 (SAR 스크리닝)

수천 개의 화합물 시트에서 특정 골격을 찾거나 리드 화합물과의 유사도를 순위 매길 때 사용합니다. 자세한 작용기 패턴 사전은 **[Excel 심화 (SAR·SDF) 가이드](/ko/guide/excel-cheminformatics)** 페이지에서 확인하세요.

| 목적 | 수식 | 결과 예시 | 활용 설명 |
| :--- | :--- | :--- | :--- |
| **특정 부분구조 포함 여부** | `=SMILES.MATCH(대상셀, "찾을구조")` | `TRUE` / `FALSE` | 대상 분자 안에 지정한 SMILES/SMARTS 골격이 들어있으면 `TRUE` 반환 |
| **두 분자 간 2D 유사도 점수** | `=SMILES.SIMILARITY(분자1, 분자2, [방식])` | `0.0` ~ `1.0` | 두 분자의 2D 지문 유사도 계산 (`"tanimoto"` 기본, `"tversky"`, `"euclid-sub"` 지원) |

---

## 5. 🔄 화학 표기법 상호 변환 (이름 · CAS · SMILES · InChI)

| 변환 방향 | 수식 | 결과 예시 |
| :--- | :--- | :--- |
| **시약 이름 / CAS 번호 ➔ SMILES** | `=NAME2SMILES("50-78-2")` 또는 `=NAME2SMILES("Aspirin")` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` / `=KET2INCHI(A2)` | `"InChI=1S/C9H8O4/c1-6(10)..."` |
| **InChI ➔ SMILES** | `=INCHI2SMILES(A2)` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` / `=INCHI2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

---

## 6. 🎨 이미지 배경색(`[background]`) & 스타일 옵션(`[style]`) 전체 규격

### 두 번째 인수: 배경색 (`[background]`)
| 입력값 | 효과 | 예시 |
| :--- | :--- | :--- |
| `""` 또는 `"trans"` / `"transparent"` | 알파 채널 투명 배경 (기본값) | `=SMILES2IMG(A2, "")` |
| `"white"` | 불투명한 흰색 사각형 배경 | `=SMILES2IMG(A2, "white")` |
| `"#RRGGBB"` 또는 CSS 색상명 | 지정한 색상으로 배경 채우기 (`"#1e1e1e"`, `"yellow"` 등) | `=SMILES2IMG(A2, "#1e1e1e", "white")` |

### 세 번째 인수: 스타일 · 라벨 · 변형 키워드 (`[style]`)
| 분류 | 대표 키워드 | 함께 쓸 수 있는 별칭(Alias) | 상세 효과 |
| :--- | :--- | :--- | :--- |
| **색상 모드** | `color` | `colour`, `cpk`, `true` | IUPAC 표준 원소별 컬러 표시 (기본값) |
| **흑백 모드** | `bw` | `mono`, `monochrome`, `black`, `gray`, `grayscale` | 학술지(ACS) 및 특허 출원용 흑백 선화 |
| **선 잉크 색상** | `white` / `ink=#HEX` | `light`, `line=#HEX`, `base=#HEX` | 다크 모드용 흰색 결합선 또는 **원하는 브랜드/슬라이드 색상**(`"ink=#1e40af"`, `"line=navy"`) 지정 |
| **부분구조 강조** | `hl:<패턴>` | `highlight:<패턴>`, `hl=<패턴>`, `highlight=<패턴>` | 지정한 SMILES/SMARTS 부분구조를 컬러로 하이라이트 |
| **원자 번호** | `num` | `number`, `numbers`, `idx`, `index`, `atom-num`, `atom-ids` | NMR/MS 피크 지정을 위한 원자 인덱스 번호(1, 2, 3...) 표시 |
| **입체화학** | `stereo` | `chiral`, `abs`, `ext`, `stereochemistry` | 카이랄 중심 입체화학 라벨(*R/S*, *E/Z*) 표시 |
| **수소 펼침** | `h` | `hydrogen`, `hydrogens`, `unfoldh`, `all-h` | 생략된 수소 원자(`H`)를 모두 펼쳐서 표시 |
| **탄소 기호** | `all-c` | `carbon`, `carbons`, `c` | 모든 탄소 골격 꼭짓점에 `C` 원소 기호 표시 |
| **대칭 반전** | `flip` / `flipy` | `flipx` (좌우 반전), `flipy` (상하 반전) | 분자 구조를 좌우(`flip`) 또는 상하(`flipy`)로 대칭 반전 |
| **회전** | `rot90` / `rot180` / `rot270` | *(4번째 `[transform]` 인수에도 입력 가능)* | 분자 구조를 시계 방향으로 90°, 180°, 270° 회전 |

---

## 🚀 다음 단계로 이어서 보기

<div class="s2i-card-grid">
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel 심화: 물성 · SAR · SDF</div>
    <p class="s2i-card-desc">12대 필수 작용기 SMARTS 패턴 사전과 HRMS m/z 부가 이온 계산 수식 보기</p>
  </a>
  <a href="../guide/faq" class="s2i-card">
    <div class="s2i-card-title">❓ 자주 묻는 질문 (FAQ)</div>
    <p class="s2i-card-desc">설치·실행, 엑셀 수식 오류 해결, 동료 PC 파일 공유 시 호환성 안내</p>
  </a>
</div>

