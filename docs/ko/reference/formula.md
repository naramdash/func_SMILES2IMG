# 엑셀 함수 및 파라미터 상세 규격

**SMILES for Microsoft Office Suite**에서 제공하는 Excel 워크시트 함수의 상세 규격 및 파라미터 매뉴얼입니다.

---

## 📐 함수 서명 (Signatures)

### 1. `SMILES2IMG`
```text
=SMILES2IMG(smiles_or_ket, [background], [options])
```
* **용도:** 셀 내부에 진짜 **인셀(In-Cell) 이미지**로 화학 구조식을 삽입합니다 (Microsoft 365 및 Excel 2024 이상).
* **반환값:** 셀 크기에 맞춰 자동으로 크기가 조절되는 인셀 이미지 객체.

### 2. `SMILES2IMG.FLOAT`
```text
=SMILES2IMG.FLOAT(smiles_or_ket, [background], [options])
```
* **용도:** 셀 좌표에 고정(`xlMoveAndSize`)되는 **플로팅 그림 도형**을 생성합니다. Excel 2016, 2019, 2021과 호환됩니다.
* **반환값:** 셀 텍스트를 반환하며 연동된 오버레이 도형을 자동 관리합니다.

### 3. `KET2IMG` 및 `KET2IMG.FLOAT`
```text
=KET2IMG(ket_json_string, [background], [options])
=KET2IMG.FLOAT(ket_json_string, [background], [options])
```
* **용도:** 반응 화살표, 플러스 기호, 촉매 및 조건 텍스트 주석이 포함된 다단계 유기 화학 반응식(KET 포맷)을 렌더링합니다.

### 4. `INCHI2IMG` 및 `INCHI2IMG.FLOAT`
```text
=INCHI2IMG(inchi_string, [background], [options], [transform])
=INCHI2IMG.FLOAT(inchi_string, [background], [options], [transform])
```
* **용도:** IUPAC 표준 InChI 문자열(`InChI=1S/...` 및 접두사 없는 `1S/...`)로부터 화학 구조식을 직접 렌더링합니다.

### 5. `NAME2IMG`, `NAME2IMG.FLOAT`, 및 `NAME2SMILES` *(v1.7.0 신규)*
```text
=NAME2IMG(name_or_cas, [background], [options], [transform])
=NAME2IMG.FLOAT(name_or_cas, [background], [options], [transform])
=NAME2SMILES(name_or_cas)
```
* **용도:** 화합물 영문명(예: `"Aspirin"`, `"Caffeine"`) 또는 CAS 등록번호(예: `"50-78-2"`)를 입력하면 표준 SMILES로 변환하거나 즉시 구조식 이미지를 렌더링합니다 (60여 종 주요 시약 내장 오프라인 사전 + PubChem REST API 및 영구 디스크 캐시 연동).

---

## 🔬 분자 화학 정보 및 약물성(ADME/Lipinski) 산출 함수

SMILES, Ketcher KET 또는 IUPAC InChI 구조식 문자열로부터 물질명, 분자량, 정밀 질량, 화학식, 물리화학적 약물성 지표를 자동 계산/조회하는 함수군입니다:

| 함수 | 반환 타입 | 기능 설명 | 반환값 예시 |
| :--- | :--- | :--- | :--- |
| `=SMILES2NAME(smiles)` / `=KET2NAME(ket)` / `=INCHI2NAME(inchi)` | `문자열` | 물질 일반명 자동 조회 (PubChem 연동 + 로컬 캐시) | `"Aspirin"`, `"Caffeine"` |
| `=NAME2SMILES(name_or_cas)` | `문자열` | 물질명 또는 CAS 번호를 Canonical SMILES로 역변환 | `"CC(=O)Oc1ccccc1C(=O)O"` |
| `=SMILES2MW(smiles)` / `=KET2MW(ket)` / `=INCHI2MW(inchi)` | `실수` | IUPAC 분자량 자동 계산 (g/mol, 순수 로컬 계산) | `180.16`, `194.19` |
| `=SMILES2EXACTMASS(smiles)` / `=KET2EXACTMASS(ket)` / `=INCHI2EXACTMASS(inchi)` | `실수` | HRMS/MS 분석용 단일동위원소 정밀 질량 (Da, 순수 로컬) | `180.0423` |
| `=SMILES2FORMULA(smiles)` / `=KET2FORMULA(ket)` / `=INCHI2FORMULA(inchi)` | `문자열` | 표준 Hill 표기법 화학식 산출 (순수 로컬 계산) | `"C9H8O4"`, `"C8H10N4O2"` |
| `=SMILES2INCHI(smiles)` / `=KET2INCHI(ket)` | `문자열` | IUPAC 표준 InChI 문자열 변환 (순수 로컬 계산) | `"InChI=1S/C9H8O4/c1-6(10)13-8-5..."` |
| `=INCHI2SMILES(inchi)` | `문자열` | InChI를 Canonical SMILES로 변환 (순수 로컬 계산) | `"CC(=O)Oc1ccccc1C(=O)O"` |
| `=SMILES2INCHIKEY(smiles)` / `=KET2INCHIKEY(ket)` / `=INCHI2INCHIKEY(inchi)` | `문자열` | 27자리 표준 InChIKey 고유 해시 생성 (순수 로컬 계산) | `"BSYNRYMUTXBXSQ-UHFFFAOYSA-N"` |
| `=SMILES2LOGP(structure)` | `실수` | Wildman-Crippen 옥탄올-물 분배계수 LogP (순수 로컬) | `1.31` |
| `=SMILES2TPSA(structure)` | `실수` | 위상 극성 표면적 TPSA (Å², 순수 로컬) | `63.60` |
| `=SMILES2HBD(structure)` | `정수` | 수소 결합 공여체 수 (OH + NH, 순수 로컬) | `1` |
| `=SMILES2HBA(structure)` | `정수` | 수소 결합 수용체 수 (N + O, 순수 로컬) | `4` |
| `=SMILES2ROTBONDS(structure)` | `정수` | 회전 가능한 단일 결합 수 (순수 로컬) | `3` |
| `=SMILES2LIPINSKI(structure)` | `문자열` | Lipinski's Rule of 5(경구 투여 약물성 규칙) 통과 여부 (`"PASS"`/`"FAIL"`) | `"PASS"` |

---

## 🧪 화학정보학(Cheminformatics) 및 SAR 스크리닝 함수 *(v1.7.0 신규)*

| 함수 | 반환 타입 | 기능 설명 | 수식 예시 |
| :--- | :--- | :--- | :--- |
| `=SMILES.MATCH(target, query)` | `논리값` | `target` 분자 내에 `query`(SMILES/SMARTS) 부분구조가 포함되어 있으면 `TRUE` 반환 | `=SMILES.MATCH(A2, "c1ccccc1")` |
| `=SMILES.SIMILARITY(mol1, mol2, [metric])` | `실수` | 두 분자 간의 2D 지문(Fingerprint) 유사도(`0.0` ~ `1.0`) 계산. `metric`: `"tanimoto"`(기본) 또는 `"tversky"` | `=SMILES.SIMILARITY(A2, $B$1, "tanimoto")` |

> 💡 **활용 팁:** 구조 기반 물리화학적 계산 및 부분구조/유사도 함수는 네트워크 요청 없이 0ms 순수 오프라인으로 작동합니다. 엑셀 필터 및 조건부 서식과 결합해 수천 개의 화합물 라이브러리를 즉시 스크리닝할 수 있습니다!

---

## ⚙️ 파라미터 규격

| 파라미터 | 데이터 타입 | 기본값 | 상세 설명 |
| :--- | :--- | :--- | :--- |
| `smiles_or_ket` | `문자열` | *(필수)* | 유효한 SMILES 문자열(예: `CN1C=NC2=C1C(=O)N(C(=O)N2C)C`) 또는 Ketcher JSON 문자열(`{"root": ...}`). |
| `background` | `문자열` | `""` (자동) | 배경색. 빈 문자열(`""`) 입력 시 기본 투명/흰색 처리. 16진수 HEX 코드 (`"#FFFFFF"`, `"#1A1A1A"`, `"#00000000"`), 색상 이름 (`"white"`, `"transparent"`, `"black"`) 지원. |
| `options` | `문자열` | `""` | 공백, 쉼표, 세미콜론, 파이프로 구분된 스타일 및 렌더링 옵션 플래그. |

---

## 🎨 옵션 키워드 상세 목록

| 옵션 키워드 | 분류 | 상세 기능 | 수식 예시 |
| :--- | :--- | :--- | :--- |
| `bw` / `monochrome` | 모드 | 특허청 및 ACS 논문용 흑백 선화 모드 | `"bw"` |
| `colored` | 모드 | IUPAC 표준 CPK 원소 색상 표기 | `"colored"` |
| `hl:<smarts>` / `highlight:<smarts>` | 강조 | 지정한 SMILES/SMARTS 부분구조(스캐폴드/작용기) 하이라이트 표시 | `"hl:c1ccccc1"` |
| `num` / `numbers` | 라벨 | NMR/MS 피크 지정을 위한 원자 번호(1-based) 표기 | `"num"` |
| `stereo` | 라벨 | CIP 입체화학 라벨 (*R/S*, *E/Z*) 강제 표기 | `"stereo"` |
| `h` / `hydrogens` | 라벨 | 생략된 수소 원자(Explicit H) 모두 펼침 표기 | `"h"` |
| `all-c` / `carbons` | 라벨 | 모든 탄소 골격에 원소 기호(`C`) 명시 | `"all-c"` |
| `white` / `ink=#HEX` | 모드 | 다크 모드용 흰색 잉크 또는 사용자 정의 잉크 색상 | `"white"` |
| `flip` | 변형 | 구조식 좌우(수평) 반전 | `"flip"` |
| `flipv` | 변형 | 구조식 상하(수직) 반전 | `"flipv"` |
| `rot90` | 변형 | 시계 방향 90도 회전 | `"rot90"` |
| `rot180` | 변형 | 180도 회전 | `"rot180"` |
| `rot270` | 변형 | 시계 방향 270도 회전 | `"rot270"` |

---

## 🔀 자유로운 구분자 문법

옵션 토큰들은 공백, 쉼표, 세미콜론, 파이프 기호 중 편한 방식을 사용할 수 있습니다:

```text
=SMILES2IMG(A2, "", "bw num stereo")
=SMILES2IMG(A2, "", "bw, num, stereo")
=SMILES2IMG(A2, "", "bw; num; stereo")
=SMILES2IMG(A2, "", "bw | num | stereo")
```

---

## 💡 실전 활용 예시

* **특정 코어 스캐폴드/작용기 하이라이트 렌더링:**
  ```text
  =SMILES2IMG(A2, "trans", "hl:c1ccccc1")
  ```
* **CAS 번호 또는 시약 이름으로 즉시 구조식 생성:**
  ```text
  =NAME2IMG("50-78-2")
  =NAME2IMG("Ibuprofen", "white")
  ```
* **특허 출원용 흑백 선화:**
  ```text
  =SMILES2IMG(A2, "", "bw")
  ```
* **다크 테마 시트용 흰색 잉크 렌더링:**
  ```text
  =SMILES2IMG(A2, "#1e1e1e", "white")
  ```
* **NMR 피크 분석용 원자 번호 & 입체화학:**
  ```text
  =SMILES2IMG(A2, "", "num stereo")
  ```
* **다단계 유기 합성 반응식 전체 삽입:**
  ```text
  =KET2IMG(B2)
  ```
