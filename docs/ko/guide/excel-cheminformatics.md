# Excel 레시피 2부: 물성 · Lipinski · SAR · SDF 연동

> **Excel 레시피 2부:** 구조식 옆 열에 분자량, HRMS 정밀질량, Lipinski 약물성, 3D 복잡도(`Fsp3`), 부분구조·유사도 SAR 스크리닝, 표기법 변환 및 12열 `.sdf` 입출력을 연결하는 실무 레시피입니다. 모든 레시피는 실제 Excel 구동 영상과 스크린샷을 포함합니다.

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf" class="s2i-card">
    <div class="s2i-card-title">📥 실습용 화합물 DB 다운로드 (.sdf)</div>
    <p class="s2i-card-desc">[SMILES] ➔ [SDF File] ➔ [Import SDF]로 열어 12열 SAR &amp; 3D 복잡도 시트 생성을 바로 체험하세요</p>
  </a>
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 8개 탭 마스터 쇼케이스 엑셀 파일 (.xlsx)</div>
    <p class="s2i-card-desc">0.빠른시작 · 1.구조식생성 · 2.물성/Lipinski/3D · 3.부분구조/유사도 · 4.스타일매트릭스 · 5.반응식 · 6.구형호환 · 7.리본도구</p>
  </a>
</div>

---

## 레시피 6 & 7: 분자량, HRMS 정밀질량, Lipinski Rule of 5 & 3D 복잡도

**목표:** 오프라인 0ms 계산으로 분자량(`MW`), 단일동위원소 정밀질량, `LogP`, `TPSA`, `HBD/HBA/RotBonds`, 고리 수, 방향족 고리 수, `Fsp3`, 키랄 중심 및 Lipinski `PASS`/`FAIL`을 계산합니다.

```text
=SMILES2MW(A2)
=SMILES2EXACTMASS(A2)
=SMILES2LIPINSKI(A2)
=SMILES2FSP3(A2)
```

<FeatureMedia id="excel-cheminformatics" />

| 물성 지표 | 수식 | Lipinski / 신약개발 기준 | 아스피린 예시 | 실무 팁 |
| :--- | :--- | :---: | :---: | :--- |
| **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | — | `C9H8O4` | 표준 Hill 표기법 |
| **평균 분자량 (MW)** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | 당량·칭량 계산용 숫자 반환 |
| **정밀 질량 (HRMS)** | `=SMILES2EXACTMASS(A2)` | — | `180.0423` | `[M+H]+`: `+1.0073`, `[M+Na]+`: `+22.9892` |
| **LogP / TPSA** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | **≤ 5.0** / **≤ 140 Å²** | `1.31` / `63.60` | CNS BBB 투과 ≤ 90 Å² |
| **HBD / HBA / RotB** | `=SMILES2HBD` / `HBA` / `ROTBONDS` | **≤ 5** / **≤ 10** / **≤ 10** | `1` / `4` / `3` | 경구 흡수 핵심 지표 |
| **Lipinski 판정** | `=SMILES2LIPINSKI(A2)` | — | `PASS` | 4대 규칙 종합 `"PASS"`/`"FAIL"` |
| **중원자 / 고리 / 방향족** | `=SMILES2HEAVY` / `RINGS` / `AROMATIC` | **방향족 ≤ 3** | `13` / `1` / `1` | 수용성 및 골격 추적 |
| **3D Fsp³ / 키랄** | `=SMILES2FSP3(A2)` / `=SMILES2CHIRAL(A2)` | **Fsp³ ≥ 0.25** | `0.11` / `0` | 입체 복잡도 ("Escape from Flatland") |

> 🔄 입력이 InChI 또는 KET인 경우 앞의 `SMILES2`를 `INCHI2` / `KET2`로 바꾸면 똑같이 계산됩니다 (`=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).

---

## 레시피 8: 부분구조 검색(`SMILES.MATCH`), 강조 렌더링 & 유사도 순위(`SAR`)

**목표:** 특정 작용기를 가진 분자만 `TRUE`/`FALSE`로 걸러내고, 해당 골격을 컬러로 강조하며, 리드 화합물과 Tanimoto 유사도(`0.0~1.0`)를 매깁니다.

```text
=SMILES.MATCH(A2, "c1ccccc1")
=SMILES.SIMILARITY(A2, $B$2)
=SMILES2IMG(A2, "", "hl:c1ccccc1")
```

<FeatureMedia id="excel-sar-screening" />

### 바로 복사해 쓰는 12대 작용기 SMARTS 패턴 사전
| 작용기 / 골격 | 패턴 코드 | 포함 여부 수식 | 하이라이트 이미지 수식 |
| :--- | :--- | :--- | :--- |
| **벤젠 고리** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **피리딘 고리** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **카복실산 (`-COOH`)** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **에스터 (`-COO-`)** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **아마이드 (`-CONH-`)** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **설폰아마이드 (`-SO₂NH-`)** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **케톤** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **알데하이드** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **페놀** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **나이트로기** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **할로겐 원소** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **트라이플루오로메틸** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

> 유사도 3번째 인수 `[metric]`은 `"tanimoto"`(기본), `"tversky"`, `"euclid-sub"`를 지원합니다.

---

## 레시피 9: 이름 · CAS · SMILES · InChI · InChIKey 상호 변환

**목표:** CAS 번호나 물질명을 표준 SMILES, InChI, 27자리 InChIKey로 변환해 시트 간 중복 제거(`VLOOKUP`/`XLOOKUP`)에 활용하고 `=INCHI2IMG`로 렌더링합니다.

```text
=NAME2SMILES("50-78-2")
=SMILES2INCHI(B2)
=SMILES2INCHIKEY(B2)
=INCHI2IMG(C2)
```

<FeatureMedia id="excel-conversions" />

---

## 레시피 10: 12열 `.SDF` / `.MOL` 화학 DB 가져오기 & 내보내기

**목표:** 외부 `.sdf` 화합물 라이브러리를 열어 12개 의약화학·3D 복잡도 열(`Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski`)을 갖춘 라이브 시트를 생성하고 다시 `.sdf`로 저장합니다.

<FeatureMedia id="excel-sdf-workflow" />

1. 리본 **`[SMILES] ➔ [SDF File] ➔ [Import SDF File (.sdf)...]`**를 클릭하고 `.sdf` 또는 `.mol` 파일을 선택합니다.
2. 12개 열의 수식과 행 높이·열 너비가 자동으로 맞춰져 생성됩니다.
3. 시트를 도킹 프로그램 등에 전달할 때는 **`[Export Sheet to SDF (.sdf)...]`**를 선택합니다.

---

## 엑셀 `IF` / `AND` & 조건부 서식 결합

| 판정 목표 | 복사해 쓸 수식 (`A2` = SMILES) |
| :--- | :--- |
| **피리딘 보유 + Lipinski 통과 히트** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "✅ 후보 선정", "제외")` |
| **CNS BBB 투과 후보 (`TPSA ≤ 90`, `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "🧠 CNS 유망", "일반")` |
| **3D 유망 리드 (`Fsp3 ≥ 0.25`, `방향족 ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "✨ 3D 유망", "평면")` |

---

## 🚀 이어서 살펴보기

<div class="s2i-card-grid">
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📑 40개 전체 수식 &amp; 스타일 사전</div>
    <p class="s2i-card-desc">40종 전체 함수 규격, 반환 데이터 타입 및 오류 코드 안내</p>
  </a>
  <a href="../tools/formula-builder" class="s2i-card">
    <div class="s2i-card-title">⚡ 인터랙티브 수식 생성기</div>
    <p class="s2i-card-desc">클릭만으로 스타일·하이라이트 옵션을 조합하고 완성된 수식 복사하기</p>
  </a>
</div>
