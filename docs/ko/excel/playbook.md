# Excel 예제 파일 (Playbook `.xlsx`)

웹 문서를 함께 열지 않아도 Microsoft Excel 안에서 49종 전체 화학 수식, 물성·3D 대시보드, SMARTS 하이라이트, KET 반응식, 실전 실험실 템플릿과 오프라인 함수 사전까지 직접 실습할 수 있는 10개 탭 공식 Playbook입니다.

<PlaybookPreview app="excel" />

::: tip 활용 방법
1. 먼저 `func_SMILES2IMG`를 설치합니다 ([설치 가이드](/ko/guide/install)).
2. <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx (.xlsx)</strong></a>를 내려받아 **Microsoft Excel**에서 바로 엽니다.
3. 아래 **목적별 실습 코스**를 따라 해 보거나, 당량 계산기·Hit-to-Lead 시트를 본인 연구 파일에 그대로 복사해 활용해 보세요.
:::

## 목적별 추천 실습 코스

| 추천 코스 | 열어볼 시트 탭 | 핵심 실습 내용 |
| :--- | :--- | :--- |
| **1. 처음 시작하는 입문 코스** | `0. Start Here` ➔ `1. Structure Rendering` | 입력 셀의 SMILES·시약명을 바꿔 보고, 구조식 셀을 **더블클릭**해 **Ketcher 2D** 제자리 수정을 체험합니다. |
| **2. 유기 합성 실험노트 (ELN) 코스** | `6. Reactions & KET` ➔ `7. Real-World Workflows` | 다단계 합성 반응식(`=KET2IMG`)을 더블클릭해 촉매 조건을 수정하고, `=SMILES2MW` 연동 당량 계산기로 시약 칭량(`mg`)과 수율(`%`)을 자동 산출합니다. |
| **3. 의약화학 SAR 스크리닝 코스** | `2. Properties & 3D` ➔ `3. SAR & SMARTS` | 쿼리 셀(`$B$3`)의 SMARTS 패턴 하나만 바꿔 `=SMILES.MATCH`, 빨간색 골격 하이라이트(`hl:`), `=INCHI.MATCH` 일치가 갱신되는 것을 확인합니다. |

## 시트별 상세 구성 (10개 탭)

| 시트 탭 | 수록된 실습 내용 및 활용법 | 대표 수식 · 기능 |
| :--- | :--- | :--- |
| **`0. Start Here`** | 체험 샌드박스(셀 수정·더블클릭 실습) 및 10개 탭 전체 학습 로드맵 | 핵심 수식 10종 & 더블클릭 편집 |
| **`1. Structure Rendering`** | 영문 화합물명, CAS 번호, SMILES, IUPAC InChI 12종 자동 2D 렌더링 | `=NAME2IMG`, `=SMILES2IMG`, `=INCHI2IMG`, `=NAME2SMILES` |
| **`2. Properties & 3D`** | 15열 신약 SAR·3D 복잡도 표 + `AVERAGE` / `COUNTIF` 실시간 KPI 요약 행 | `=SMILES2FORMULA`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2TPSA`, `=SMILES2FSP3` |
| **`3. SAR & SMARTS`** | 쿼리 셀(`$B$3`) 수정 시 즉시 반응하는 하이라이트, `INCHI.MATCH` & 12종 SMARTS 사전 | `=SMILES.MATCH`, `=INCHI.MATCH`, `=SMILES2IMG(..., "hl:...")` |
| **`4. Notation Conversions`** | SMILES·InChI·InChIKey 상호 변환 및 `InChIKey` 기반 중복 화합물 자동 검출 | `=SMILES2INCHI`, `=SMILES2INCHIKEY`, `=INCHI2SMILES`, `=INCHI2MW` |
| **`5. Styles & Transforms`** | ACS 흑백, 다크 모드, 커스텀 잉크, NMR 번호, 입체표시, 평면 회전 16종 비교 매트릭스 | `"bw"`, `"white"`, `"ink=#1E40AF"`, `"num"`, `"stereo"`, `"rot90"` |
| **`6. Reactions & KET`** | 실제 아스피린·파라세타몰 합성 반응식 및 단일 분자 무손실 KET 캔버스 | `=KET2IMG`, `=KET2SMILES`, `=KET2FORMULA`, `=KET2MW`, 더블클릭 Ketcher 편집 |
| **`7. Real-World Workflows`** | 바로 복사해 쓰는 유기 합성 당량 스케일업 계산기(`mg`·수율 `%`) & Hit-to-Lead MPO 선별표 | `=SMILES2MW(C6)*H6` 자동 환산 & `IF(AND(...))` 후보 판정 |
| **`8. Legacy Excel (.FLOAT)`** | Excel 2016, 2019, 2021 환경을 위한 셀 고정 부동 그림(`.FLOAT`) 호환 수식 10종 | `=SMILES2IMG.FLOAT`, `=NAME2IMG.FLOAT` |
| **`9. Ribbon, SDF & Ref`** | `[SMILES]` 리본·3단계 우클릭 메뉴표, `.sdf` 11열 가져오기 가이드, 49종 전체 함수 및 오류 해결 | 오프라인 통합 사전 (`#N/A`, `#VALUE!`, `#NAME?`) |

## 관련 가이드

- [**49개 함수 전체 개요**](/ko/excel/formulas) — 전체 49종 함수 규격, 반환 타입 및 공용 스타일 키워드 사전
- [**SMILES · InChI**](/ko/excel/structures) · [**시약명 · 용매 · CAS**](/ko/excel/name-cas) — 내 엑셀 시트에서 구조식 표와 이름 기반 렌더링 시작하기
- [**3종 Office 예제 파일 허브**](/ko/guide/playbooks) — [Word (.docx)](/ko/word/playbook) · [PowerPoint (.pptx)](/ko/powerpoint/playbook) 예제 파일 비교하기
