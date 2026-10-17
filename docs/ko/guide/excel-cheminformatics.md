# Excel 심화: 물성 계산 · 신약 스크리닝(SAR) · SDF 연동

엑셀에서 구조식 그림을 그리는 것을 넘어, **분자량·정밀질량·화학식 자동 계산**, **경구 약물성(Lipinski Rule of 5) 판정**, **부분구조 검색 및 2D 유사도 스크리닝**, 그리고 **`.sdf` 화학 데이터베이스 파일 입출력**까지 활용하는 실무 가이드입니다.

::: tip ⚡ 100% 오프라인 · 대기 시간 0초
물질명 웹 조회(`=SMILES2NAME`)를 제외한 모든 물리화학 계산과 부분구조·유사도 검색은 외부 서버 통신 없이 **내 PC 안에서 0ms 즉시 계산**됩니다. 수천 개의 화합물이 담긴 시트에서도 필터와 정렬이 실시간으로 작동합니다.
:::

---

## 1. 기본 분자 물성 자동 계산 (이름 · 화학식 · 분자량 · 정밀질량)

A2 셀에 SMILES(예: 아스피린 `CC(=O)Oc1ccccc1C(=O)O`)가 들어있을 때 바로 옆 열에 입력할 수 있는 기본 물성 함수입니다:

| 구하고 싶은 항목 | 입력할 수식 | 결과값 (아스피린 예시) | 실무 활용 팁 |
| :--- | :--- | :--- | :--- |
| **물질 일반명 (Name)** | `=SMILES2NAME(A2)` | `"Aspirin"` | 내장 60종 주요 시약 사전 + PubChem 자동 조회 및 영구 캐시 |
| **이름/CAS ➔ SMILES** | `=NAME2SMILES("50-78-2")` | `"CC(=O)Oc1ccccc1C(=O)O"` | 시약 이름이나 CAS 등록번호만 있을 때 표준 SMILES 추출 |
| **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | `"C9H8O4"` | 탄소(C), 수소(H), 알파벳 순의 표준 Hill 표기법 화학식 |
| **평균 분자량 (MW)** | `=SMILES2MW(A2)` | `180.16` | 시약 당량(eq) 및 몰 농도(M) 계산용 표준 분자량 (g/mol) |
| **정밀 질량 (Exact Mass)** | `=SMILES2EXACTMASS(A2)` | `180.0423` | **HRMS / LC-MS 질량분석 스펙트럼 피크 확인용** 단일동위원소 질량 (Da) |
| **표준 InChI / InChIKey** | `=SMILES2INCHIKEY(A2)` | `"BSYNRYMUTXBXSQ-..."` | 서로 다른 데이터베이스 간 중복 화합물 제거(중복 값 찾기)에 최적 |

::: info 🧪 LC-MS / HRMS 연구자를 위한 m/z 부가 이온(Adduct) 계산 팁
질량분석기에서 관찰되는 피크는 평균 분자량(`MW`)이 아니라 가장 풍부한 동위원소(¹²C, ¹H, ¹⁶O)만으로 계산된 **`=SMILES2EXACTMASS(A2)`** 값을 기준으로 해야 정확합니다:
* **`[M+H]⁺` (양성자 부가):** `=SMILES2EXACTMASS(A2) + 1.0073`
* **`[M+Na]⁺` (소듐 부가):** `=SMILES2EXACTMASS(A2) + 22.9892`
* **`[M−H]⁻` (양성자 탈리):** `=SMILES2EXACTMASS(A2) - 1.0073`
:::

> 🔄 **입력 데이터가 InChI 또는 반응식(KET)인 경우:** 함수 이름 앞의 `SMILES2` 대신 `INCHI2` 또는 `KET2`를 쓰면 똑같이 작동합니다 (예: `=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).

---

## 2. 경구 약물성(Lipinski Rule of 5) & ADME 물리화학 지표

합성하거나 구매할 후보 물질이 경구 투여 약물로서 적합한 물리화학적 성질을 가졌는지 엑셀에서 즉시 판별할 수 있습니다.

### Lipinski's Rule of 5 자동 판정 (`=SMILES2LIPINSKI`)
```text
=SMILES2LIPINSKI(A2)
```
아래 4가지 기준을 평가하여 통과 시 **`"PASS"`**, 미달 시 **`"FAIL"`** 문자열을 반환합니다:

| 물리화학 지표 | 개별 계산 수식 | Lipinski 기준 | 아스피린 예시 | 신약 개발에서의 의미 |
| :--- | :--- | :---: | :---: | :--- |
| **분자량 (MW)** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | 세포막 투과 및 장내 흡수 한계 크기 |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | **≤ 5.0** | `1.31` | 옥탄올/물 분배계수 (수용성과 지용성의 균형) |
| **수소 결합 공여체 (HBD)** | `=SMILES2HBD(A2)` | **≤ 5** 개 | `1` | `-OH`, `-NH` 그룹의 수소 개수 |
| **수소 결합 수용체 (HBA)** | `=SMILES2HBA(A2)` | **≤ 10** 개 | `4` | 질소(`N`), 산소(`O`) 원자의 개수 |

### 추가 ADME 핵심 지표: 극성 표면적(`TPSA`) & 회전 결합 수(`ROTBONDS`)
| 지표 | 수식 | 결과 예시 | 실무 해석 가이드 |
| :--- | :--- | :---: | :--- |
| **위상 극성 표면적 (TPSA)** | `=SMILES2TPSA(A2)` | `63.60` (Å²) | • **≤ 140 Å²:** 장내 경구 흡수 양호<br>• **≤ 90 Å²:** 중추신경계(CNS) 약물의 뇌혈관장벽(BBB) 투과 유리 |
| **회전 가능한 결합 수** | `=SMILES2ROTBONDS(A2)` | `3` 개 | • **≤ 10 개 (Veber 규칙):** 분자 유연성이 적절하여 경구 생체이용률 유리 |

---

## 3. 부분구조 검색(`=SMILES.MATCH`) & 필수 작용기 SMARTS 사전

수천 개의 화합물 목록에서 특정 골격(Scaffold)이나 작용기(Functional Group)가 들어있는 분자만 골라낼 때 사용합니다:

```text
=SMILES.MATCH(A2, "찾을_구조_패턴")
```
* 포함되어 있으면 **`TRUE`**, 없으면 **`FALSE`** 논리값을 반환하므로 엑셀 **자동 필터**로 원하는 계열의 화합물만 1초 만에 추려낼 수 있습니다.
* 동일한 패턴을 `=SMILES2IMG(A2, "", "hl:패턴")`에 넣으면 해당 작용기만 컬러로 하이라이트해 그림으로 보여줍니다!

### 📋 복사해서 바로 쓰는 12대 필수 작용기·골격 패턴 사전
SMARTS 문법을 외울 필요 없이 아래 표에서 원하는 작용기 코드를 복사해 `"찾을_구조_패턴"` 자리에 붙여넣으세요:

| 찾고 싶은 골격 / 작용기 | 패턴 코드 (복사용) | 검색 수식 예시 | 하이라이트 렌더링 예시 |
| :--- | :--- | :--- | :--- |
| **벤젠 고리 (방향족)** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **피리딘 고리 (헤테로방향족)** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **카복실산 (`-COOH`)** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **에스터 (`-COO-`)** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **아마이드 (`-CONH-`)** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **설폰아마이드 (`-SO₂NH-`)** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **케톤 (`-C(=O)-`)** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **알데하이드 (`-CHO`)** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **페놀 (방향족 `-OH`)** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **나이트로기 (`-NO₂`)** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **할로겐 원소 (`F, Cl, Br, I`)** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **트라이플루오로메틸 (`-CF₃`)** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

---

## 4. 리드 화합물 대비 2D 유사도 랭킹 (`=SMILES.SIMILARITY`)

기준이 되는 활성 물질(Lead compound, 예: `B1` 셀에 고정)과 시트의 각 후보 물질(`A2`, `A3`...)이 구조적으로 얼마나 유사한지 **`0.0` ~ `1.0` 사이의 점수**로 계산합니다:

```text
=SMILES.SIMILARITY(A2, $B$1)
```
> 💡 **절대참조(`$B$1`) 팁:** 두 번째 인수를 `$B$1`로 고정해 두고 수식을 아래로 드래그한 뒤, 엑셀에서 **내림차순 정렬**을 누르면 리드 물질과 가장 유사한 유도체(Analog) 순서대로 즉시 정렬됩니다!

### 3가지 유사도 계산 방식 (`[metric]` 선택 인수)
세 번째 인수로 계산 알고리즘을 선택할 수 있습니다 (생략 시 `"tanimoto"` 기본 적용):

| 메트릭 옵션 | 수식 예시 | 특징 및 추천 용도 |
| :--- | :--- | :--- |
| **`"tanimoto"`** *(기본값)* | `=SMILES.SIMILARITY(A2, $B$1, "tanimoto")` | 가장 널리 쓰이는 표준 2D 분자 지문(Fingerprint) 전체 유사도 비교 |
| **`"tversky"`** | `=SMILES.SIMILARITY(A2, $B$1, "tversky")` | 두 분자의 크기 차이가 클 때(작은 조각이 큰 분자에 포함된 정도를 볼 때) 유리한 비대칭 유사도 |
| **`"euclid-sub"`** | `=SMILES.SIMILARITY(A2, $B$1, "euclid-sub")` | 첫 번째 분자의 부분구조 특성이 두 번째 분자에 얼마나 반영되어 있는지 평가 |

---

## 5. 대용량 `.sdf` / `.mol` 화학 데이터 파일 가져오기 & 내보내기

외부 화학 데이터베이스(PubChem, ChEMBL, 사내 화합물 라이브러리 등)에서 받은 **`.sdf` 또는 `.mol` 파일**을 변환 프로그램 없이 엑셀에서 바로 열고 저장할 수 있습니다.

### 📥 `.sdf` 파일을 8열 약물화학 대시보드 표로 불러오기 (`Import SDF File (.sdf)...`)
1. 엑셀 상단 **`[SMILES]`** 리본 탭에서 **`[SDF File]` ➔ `[Import SDF File (.sdf)...]`** 메뉴를 클릭합니다.
2. 내 PC에서 `.sdf` 또는 `.mol` 파일을 선택하면 현재 시트에 아래와 같은 **8열 표준 약물화학 분석 표**가 행 높이·열 너비까지 맞춰져 자동으로 생성됩니다:

| 열 (Column) | 헤더 이름 | 자동으로 채워지는 내용 (`2행` 기준) |
| :---: | :--- | :--- |
| **A열** | **Name** | `.sdf` 파일 내 화합물 이름 또는 `<NAME>` / `<TITLE>` 태그 |
| **B열** | **Structure** | `=SMILES2IMG(C2)` *(구버전 엑셀에서는 자동으로 `=SMILES2IMG.FLOAT(C2)` 입력)* |
| **C열** | **SMILES** | 정규화된 표준 Canonical SMILES 문자열 |
| **D열** | **Formula** | `=SMILES2FORMULA(C2)` (표준 Hill 화학식) |
| **E열** | **MW (g/mol)** | `=SMILES2MW(C2)` (평균 분자량) |
| **F열** | **Exact Mass** | `=SMILES2EXACTMASS(C2)` (HRMS 정밀 질량) |
| **G열** | **LogP** | `=SMILES2LOGP(C2)` (지용성 분배계수) |
| **H열** | **Lipinski** | `=SMILES2LIPINSKI(C2)` (`"PASS"` / `"FAIL"` 약물성 판정) |

### 📤 엑셀 시트를 `.sdf` 파일로 내보내기 (`Export Worksheet to SDF...`)
1. 상단 **`[SMILES]`** ➔ **`[SDF File]`** ➔ **`[Export Worksheet to SDF...]`** 메뉴를 클릭하고 저장할 파일 이름을 지정합니다.
2. 복잡한 열 지정 없이도 현재 시트의 각 행에서 화합물 이름과 SMILES(또는 구조식 수식)를 자동으로 찾아 표준 V2000 `.sdf` 파일로 묶어 내보내므로, 분자 도킹(Docking) 프로그램이나 타 연구실에 즉시 전달할 수 있습니다.

---

## 6. 엑셀 기본 함수(`IF`, `FILTER`) 및 조건부 서식 결합 레시피

모든 화학 수식은 엑셀의 표준 숫자(`Number`), 논리값(`TRUE`/`FALSE`), 텍스트를 반환하므로 엑셀 기본 함수 및 **조건부 서식**과 완벽하게 호환됩니다.

### ① 다중 조건 히트(Hit) 화합물 자동 판정 (`=IF` + `AND`)
| 판정 목적 | 바로 복사해 쓸 엑셀 수식 (`A2` = SMILES) |
| :--- | :--- |
| **Lipinski 통과 + 특정 코어(`c1ccncc1`) 보유 시 `"후보 선정"`** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "✅ 후보 선정", "제외")` |
| **뇌혈관장벽(BBB) 투과 유망 물질 판정 (`TPSA ≤ 90` & `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "🧠 CNS 유망", "일반")` |
| **할로겐 원소 보유 여부에 따라 구조식 하이라이트 자동 전환** | `=IF(SMILES.MATCH(A2, "[F,Cl,Br,I]"), SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]"), SMILES2IMG(A2))` |

### ② 조건부 서식으로 `PASS` / `FAIL` 및 유사도 막대 그래프 입히기 (10초 설정)
1. **`PASS` / `FAIL` 색상 강조:** `=SMILES2LIPINSKI(A2)`가 입력된 열 전체를 선택 ➔ 엑셀 상단 **[홈] ➔ [조건부 서식] ➔ [셀 강조 규칙] ➔ [같음...]** ➔ `PASS` 입력 후 **'녹색 채우기'** 선택!
2. **유사도 점수(`0.0~1.0`) 막대 그래프:** `=SMILES.SIMILARITY(A2, $B$1)`가 입력된 열 전체를 선택 ➔ **[홈] ➔ [조건부 서식] ➔ [데이터 막대]** 클릭! 셀 안에 유사도 비율이 직관적인 게이지 막대로 표시됩니다.

---

## 🚀 다음 단계로 이어서 보기

<div class="s2i-card-grid">
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📐 전체 엑셀 함수 치트시트</div>
    <p class="s2i-card-desc">25종 이상의 전체 함수 규격, 반환 데이터 타입 및 오류 코드 사전 확인하기</p>
  </a>
  <a href="../tools/formula-builder" class="s2i-card">
    <div class="s2i-card-title">🛠️ 엑셀 함수 표현 생성기</div>
    <p class="s2i-card-desc">마우스 클릭만으로 스타일·하이라이트 옵션을 조합하고 완성된 수식 복사하기</p>
  </a>
</div>


