# 40개 엑셀 함수 전체 개요

**`func_SMILES2IMG`** 의 40개 엑셀 함수는 **`=<입력>2<출력>`** 통합 규칙으로 설계되었습니다. 입력 셀 형식(`SMILES`, `NAME`, `INCHI`, `KET`)에 맞춰 접두어를 고르고, 셀에 얻고 싶은 결과(2D 구조식 이미지, 정량 물성, SAR 판별값, 변환 식별자)에 맞춰 접미어를 선택하면 됩니다.

::: tip 40개 함수를 지배하는 2가지 공용 파이프라인
* **2D 구조식 이미지 출력 (`8종`):** `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG` (및 `.FLOAT` 4종)는 모두 동일한 4인수 규격 `(입력, [배경색], [스타일], [변형])` 과 동일한 EPAM Indigo 렌더러를 사용합니다.
* **단일 값 출력 (`32종`):** 물성·ADME·SAR·식별자 변환 함수는 엑셀 기본 실수·정수·`TRUE`/`FALSE`·문자열을 반환하므로 `IF`, `FILTER`, `SORT`, `XLOOKUP`, 당량 계산(`MW × mmol`)에 바로 결합됩니다. 모든 `SMILES2*` 물성 함수는 `InChI` 와 단일 분자 `KET` 셀도 자동으로 인식합니다.
:::

## ⭐ 실무 필수 수식 Top 5

**전체 복사** 를 클릭해 엑셀 **`A1`** 셀에 붙여넣으면 입력값(`A2:A5`)과 수식(`B2:B6`)이 한 번에 채워지며, 행별로 복사할 수도 있습니다 (구버전 Excel 2016~2021에서는 `...2IMG` 뒤에 **`.FLOAT`** 를 붙이세요):

<FormulaTable
  title="실무 필수 수식 Top 5 시트"
  :columns='["SMILES / 입력값", "수식"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG(A2)", desc: "A2 SMILES로 표준 컬러 구조식 렌더링", out: "에탄올 2D" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2IMG(A3, \"\", \"bw\")", desc: "A3 기준 논문 투고용 ACS 흑백 선화", out: "흑백 아스피린" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A4)", desc: "평균 분자량 (0ms 오프라인 실수)", out: "180.16" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LIPINSKI(A5)", desc: "Lipinski Rule of 5 약물성 판정", out: "PASS" },
    { input: "", formula: "=NAME2IMG(\"Aspirin\")", desc: "화합물 이름·CAS 단독 2D 구조식", out: "아스피린 2D" }
  ]'
/>

## 🗺️ 사이드바 섹션별 안내

| 사이드바 섹션 | 포함 함수 (`40개`) | 세부 가이드 |
| :--- | :--- | :--- |
| **구조식 이미지** | `SMILES2IMG`, `NAME2IMG`, `INCHI2IMG`, `KET2IMG` (+ `.FLOAT` 4종) | [SMILES · InChI](/ko/excel/structures) · [시약명 · 용매 · CAS](/ko/excel/name-cas) · [합성 반응식 (KET)](/ko/excel/reactions) · [스타일 · 하이라이트](/ko/excel/styles-reactions) |
| **물성 · 약물성** | `*2FORMULA`, `*2MW`, `*2EXACTMASS`, `*2NAME` (각 3종) + 11개 ADME/3D 지표 | [화학식 · 분자량 · 정밀질량](/ko/excel/properties) · [Lipinski · ADME · 3D](/ko/excel/lipinski-adme) |
| **SAR · 식별자** | `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `NAME2SMILES`, `*2INCHI`, `INCHI2SMILES`, `*2INCHIKEY` | [부분구조 · 유사도 검색](/ko/excel/sar) · [표기법 변환 · InChIKey](/ko/excel/data-sdf) |
| **시트 도구 · 3D** | `.sdf` 다중 화합물 입출력, `[Batch Style]`, 17열 TSV, 3D `Mol*` | [SDF 파일 입출력](/ko/excel/sdf) · [일괄 스타일 · TSV · 3D](/ko/excel/tools-export) · [리본 · 우클릭 메뉴](/ko/excel/ribbon) |

## 1. 구조식 이미지 함수 (`8종`)

4가지 입력 형식 모두 `(입력, [배경색], [스타일], [변형])` 인수를 똑같이 사용합니다 ([SMILES · InChI](/ko/excel/structures) · [시약명 · 용매 · CAS](/ko/excel/name-cas) · [합성 반응식](/ko/excel/reactions)):

| 입력 형식 (`1번째 인수`) | Microsoft 365 / Excel 2024+ | Excel 2016~2021용 | 예시 |
| :--- | :--- | :--- | :--- |
| **SMILES** *(KET JSON 자동 감지)* | `=SMILES2IMG(셀, [배경색], [옵션], [변형])` | `=SMILES2IMG.FLOAT(...)` | `=SMILES2IMG(A2)` |
| **물질명 · CAS 번호** | `=NAME2IMG(이름_또는_CAS, [배경색], [옵션], [변형])` | `=NAME2IMG.FLOAT(...)` | `=NAME2IMG("Aspirin")` |
| **IUPAC InChI** | `=INCHI2IMG(셀, [배경색], [옵션], [변형])` | `=INCHI2IMG.FLOAT(...)` | `=INCHI2IMG(A2)` |
| **합성 반응식 · 캔버스 (KET)** | `=KET2IMG(셀, [배경색], [옵션], [변형])` | `=KET2IMG.FLOAT(...)` | `=KET2IMG(A2)` |

::: tip 빠른 단축 문법
* **최소 입력:** 첫 번째 인수만 입력하면 투명 배경 + 표준 컬러로 렌더링됩니다 (`=SMILES2IMG(A2)`).
* **2인수 흑백 단축형:** 두 번째 인수 자리에 바로 `"bw"`(또는 `FALSE`)를 적으면 투명 배경 흑백 선화가 됩니다 (`=SMILES2IMG(A2, "bw")`).
* **키워드 구분자:** 세 번째 `[스타일]` 인수 안에서 공백, 쉼표(`,`), `|`, `;`, `+` 중 어떤 기호로도 여러 키워드를 조합할 수 있습니다.
:::

## 2. 화학식 · 분자량 · 정밀질량 · 물질명 (`12종`)

물질명 조회(`*2NAME`, 1회 조회 후 로컬 캐시)를 제외한 모든 계산은 **0ms 순수 오프라인**으로 동작합니다 ([화학식 · 분자량 · 정밀질량](/ko/excel/properties)):

| 출력값 | SMILES 입력 | InChI / KET 입력 시 | 반환 타입 | 아스피린 (`A2`) |
| :--- | :--- | :--- | :---: | :--- |
| **물질 일반명** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `String` | `"Aspirin"` |
| **화학식 (Hill)** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)` · `=KET2FORMULA(A2)` | `String` | `"C9H8O4"` |
| **분자량 (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)` · `=KET2MW(A2)` | `Double` | `180.16` |
| **정밀 질량 (Exact Mass)** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)` · `=KET2EXACTMASS(A2)` | `Double` | `180.0423` |

## 3. Lipinski · ADME · 3D 입체 지표 (`11종`)

SMILES, InChI, 단일 분자 KET 셀을 모두 자동 파싱합니다 ([Lipinski · ADME · 3D](/ko/excel/lipinski-adme)):

| 지표 | 수식 | 반환 타입 | 아스피린 | 권장 기준 |
| :--- | :--- | :---: | :---: | :--- |
| **Lipinski 판정** | `=SMILES2LIPINSKI(A2)` | `String` | `"PASS"` | MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10 |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | `Double` | `1.31` | ≤ 5.0 권장 |
| **극성 표면적 (TPSA)** | `=SMILES2TPSA(A2)` | `Double` | `63.60` | ≤ 140 Å² 경구 흡수; ≤ 90 Å² CNS BBB |
| **수소 결합 공여체** | `=SMILES2HBD(A2)` | `Integer` | `1` | ≤ 5 (`-OH`, `-NH`) |
| **수소 결합 수용체** | `=SMILES2HBA(A2)` | `Integer` | `4` | ≤ 10 (`N`, `O` 원자 수) |
| **회전 가능 결합** | `=SMILES2ROTBONDS(A2)` | `Integer` | `3` | ≤ 10 권장 |
| **중원자 수** | `=SMILES2HEAVY(A2)` | `Integer` | `13` | 비수소 원자 개수 |
| **전체 고리 수 (SSSR)** | `=SMILES2RINGS(A2)` | `Integer` | `1` | 최소 고리 집합 전체 수 |
| **방향족 고리 수** | `=SMILES2AROMATIC(A2)` | `Integer` | `1` | ≤ 3 권장 |
| **입체 탄소 비율 (Fsp³)** | `=SMILES2FSP3(A2)` | `Double` | `0.11` | ≥ 0.25 (`0.00~1.00`) |
| **카이랄 중심 수** | `=SMILES2CHIRAL(A2)` | `Integer` | `0` | 정의된 사면체 입체 중심 수 |

## 4. 부분구조 · 유사도 검색 (`2종`)

두 함수 모두 SMILES, InChI, KET 입력을 지원합니다 ([부분구조 · 유사도 검색](/ko/excel/sar)):

| 출력값 | 수식 | 반환 타입 | 결과 | 동작 설명 |
| :--- | :--- | :---: | :---: | :--- |
| **부분구조 포함 여부** | `=SMILES.MATCH(대상, "쿼리")` | `Boolean` | `TRUE` / `FALSE` | 대상 분자에 SMILES/SMARTS 패턴이 포함되면 `TRUE` |
| **2D 지문 유사도** | `=SMILES.SIMILARITY(분자1, 분자2, [방식])` | `Double` | `0.0~1.0` | 2D 유사도 점수 (`"tanimoto"`, `"tversky"`, `"euclid-sub"`) |

## 5. 표기법 변환 (`7종`)

시약명, CAS 번호, 표준 SMILES, InChI, 27자리 InChIKey 간 상호 변환 함수입니다 ([표기법 변환 · InChIKey](/ko/excel/data-sdf)):

| 변환 방향 | 수식 | 반환 타입 | 결과 예시 |
| :--- | :--- | :---: | :--- |
| **시약명 / CAS ➔ SMILES** | `=NAME2SMILES("50-78-2")` / `=NAME2SMILES("Aspirin")` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / KET ➔ InChI** | `=SMILES2INCHI(A2)` · `=KET2INCHI(A2)` | `String` | `"InChI=1S/C9H8O4/..."` |
| **InChI ➔ 표준 SMILES** | `=INCHI2SMILES(A2)` | `String` | `"CC(=O)Oc1ccccc1C(=O)O"` |
| **SMILES / InChI / KET ➔ InChIKey** | `=SMILES2INCHIKEY(A2)` · `=INCHI2INCHIKEY(A2)` · `=KET2INCHIKEY(A2)` | `String` | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |

## 6. 공용 배경색(`[bg]`) · 스타일(`[style]`) 규격

8개 `...2IMG` 및 `.FLOAT` 함수가 모두 동일하게 해석하는 옵션 키워드입니다 ([스타일 · 하이라이트](/ko/excel/styles-reactions)):

<FeatureMedia id="excel-style-matrix" />

| 항목 | 키워드 / 값 | 별칭 (Alias) | 효과 |
| :--- | :--- | :--- | :--- |
| **배경색 `[bg]`** | `""` / `"trans"` · `"white"` · `#RRGGBB` | `"transparent"`, CSS 색상명 | 투명 배경, 불투명 흰색 박스, 임의 색상 채우기 |
| **컬러 / 흑백** | `color` · `bw` | `cpk`, `mono`, `black`, `grayscale` | IUPAC 원소 컬러 또는 논문 투고용 ACS 흑백 선화 |
| **선 잉크 색상** | `white` · `ink=#HEX` | `light`, `line=#HEX`, `base=#HEX` | 다크 모드 흰색 선 또는 커스텀 hex/CSS 선 색상 |
| **부분구조 강조** | `hl:<smarts>` | `highlight:<smarts>`, `hl=<smarts>` | 일치하는 SMILES/SMARTS 골격을 컬러로 강조 |
| **원자 라벨** | `num` · `stereo` · `h` · `all-c` | `idx`, `chiral`, `unfoldh`, `carbon` | 원자 번호, CIP 입체 표시, 수소 펼침, 탄소 `C` 표기 |
| **대칭 · 회전** | `flip` · `flipy` · `rot90`/`180`/`270` | `flipx` *(4번째 인수에도 사용 가능)* | 좌우/상하 반전 및 시계 방향 회전 |

## 7. 상태 및 오류 코드 안내

| 코드 | 원인 | 해결 방법 |
| :---: | :--- | :--- |
| **`#N/A`** *(0.1초)* | 비동기 렌더링 또는 PubChem 최초 조회 중입니다. | 잠시 기다리면 자동으로 그림·값으로 바뀝니다. |
| **`#VALUE!`** | 빈 셀 참조 또는 잘못된 SMILES/InChI 문법입니다. | **`[SMILES]` ➔ `[Diagnostics]`** 에서 상세 원인을 확인하세요. |
| **`#NAME?`** | 추가 기능 미설치 또는 Excel 2016~2021 환경입니다. | [설치 가이드](/ko/guide/install)를 확인하거나 **`.FLOAT`** 를 붙이세요. |
