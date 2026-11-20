# Excel 예제 파일 (Playbook)

웹 문서를 함께 열지 않아도 Microsoft Excel 안에서 40종 전체 화학 수식, 신약 물성·Lipinski 대시보드, SMARTS 하이라이트, KET 반응식, 실전 실험실 템플릿과 오프라인 함수 사전까지 직접 실습할 수 있는 10개 탭 공식 Playbook입니다.

<PlaybookPreview app="excel" />

::: tip 활용 방법
1. 먼저 `func_SMILES2IMG`를 설치합니다 ([설치 가이드](/ko/guide/install)).
2. <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx (.xlsx)</strong></a>를 내려받아 **Microsoft Excel**에서 바로 엽니다.
3. **`0. Start Here`** 탭부터 시작해 SMILES·SMARTS 입력 셀(`$B$3`)을 수정하거나, 구조식 셀을 **더블클릭**해 **Ketcher 2D** 편집을 체험하고, 당량 계산기·Hit-to-Lead 시트를 본인 업무 파일에 그대로 복사해 활용해 보세요.
:::

## 시트별 상세 구성 (10개 탭)

| 시트 탭 | 수록된 실습 내용 및 활용법 | 대표 수식 · 기능 |
| :--- | :--- | :--- |
| **`0. Start Here`** | 3분 체험 샌드박스(셀 수정·더블클릭 실습) 및 10개 탭 전체 학습 로드맵 | 핵심 수식 10종 & 단축키 |
| **`1. Structure Rendering`** | 영문 화합물명, CAS 번호, SMILES, IUPAC InChI 12종 자동 2D 렌더링 | `=NAME2IMG`, `=SMILES2IMG`, `=INCHI2IMG`, `=NAME2SMILES` |
| **`2. Properties & Lipinski`** | 18열 신약 SAR·3D 복잡도 표 + `AVERAGE` / `COUNTIF` 실시간 KPI 요약 행 | `=SMILES2FORMULA`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2HEAVY`, `=SMILES2FSP3`, `=SMILES2LIPINSKI` |
| **`3. SAR & Similarity`** | 쿼리 셀(`$B$3`) 수정 시 즉시 반응하는 하이라이트, Tanimoto 순위 & 12종 SMARTS 사전 | `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2IMG(..., "hl:...")` |
| **`4. Notation Conversions`** | SMILES·InChI·InChIKey 상호 변환 및 `InChIKey` 기반 중복 화합물 자동 검출 | `=SMILES2INCHI`, `=SMILES2INCHIKEY`, `=INCHI2SMILES`, `=INCHI2MW` |
| **`5. Styles & Transforms`** | ACS 흑백, 다크 모드, 커스텀 잉크, NMR 번호, 입체표시, 회전/반전 16종 비교 매트릭스 | `"bw"`, `"white"`, `"ink=#1E40AF"`, `"num"`, `"stereo"`, `"flip rot90"` |
| **`6. Reactions & KET`** | 실제 아스피린·파라세타몰 합성 반응식 및 단일 분자 무손실 KET 캔버스 | `=KET2IMG`, `=KET2FORMULA`, `=KET2MW`, 더블클릭 Ketcher 편집 |
| **`7. Real-World Workflows`** | 바로 복사해 쓰는 유기 합성 당량 스케일업 계산기(`mg`·수율 `%`) & Hit-to-Lead MPO 선별표 | `=SMILES2MW(C6)*H6` 자동 환산 & `IF(AND(...))` 후보 판정 |
| **`8. Legacy Excel (.FLOAT)`** | Excel 2016, 2019, 2021 환경을 위한 셀 고정 부동 그림(`.FLOAT`) 호환 수식 10종 | `=SMILES2IMG.FLOAT`, `=NAME2IMG.FLOAT` |
| **`9. Ribbon, SDF & Ref`** | `[SMILES]` 리본·3단계 우클릭 메뉴표, `.sdf` 12열 가이드, 40종 전체 함수 및 오류 해결 | 오프라인 통합 사전 (`#N/A`, `#VALUE!`, `#NAME?`) |

## 다른 Office 예제 파일 보기

3종 오피스 예제 파일 전체 요약은 [예제 파일 (Playbooks)](/ko/guide/playbooks)에서 확인할 수 있으며, [Word 예제 파일](/ko/word/playbook)이나 [PowerPoint 예제 파일](/ko/powerpoint/playbook)로 바로 이동할 수도 있습니다.
