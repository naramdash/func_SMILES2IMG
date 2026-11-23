# 합성 반응식 (KET)

반응 화살표(`➔`), `+` 기호, 촉매·용매·온도 조건 및 원자 배치 좌표까지 포함된 **다단계 유기 합성 반응식**을 `=KET2IMG` 수식으로 엑셀 셀에 저장하고, 언제든 **Ketcher 2D** 로 다시 열어 무손실 편집할 수 있습니다.

<FormulaTable
  title="합성 반응식 (KET) 시트"
  :columns='["반응식 / KET", "수식"]'
  :rows='[
    { input: "CC(=O)O.OCC>>CC(=O)OCC", formula: "=KET2IMG(A2)", desc: "유기 합성 반응식 표준 컬러 렌더링", out: "컬러 반응식" },
    { input: "CC(=O)O.OCC>>CC(=O)OCC", formula: "=KET2IMG(A3, \"\", \"bw\")", desc: "논문 투고용 ACS 흑백 반응식 렌더링", out: "흑백 반응식" }
  ]'
/>

<FeatureMedia id="excel-reactions-ket" />

## 1. 반응식 삽입 및 무손실 재편집 방법

1. 빈 셀을 선택하고 상단 리본에서 **`[SMILES]` ➔ `[Insert / Edit (KET)]`** 를 클릭합니다. *(참고: **`[Insert / Edit (SMILES)]`** 에서 화살표나 다중 분자를 그려도 자동으로 `=KET2IMG(...)` 수식으로 저장됩니다.)*
2. 오프라인 **Ketcher 2D 편집기**에서 반응물, `+`, 반응 화살표(`➔`), 화살표 위·아래의 시약/촉매/온도 텍스트, 생성물을 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd> 를 누르면 `=KET2IMG("...")` (Excel 2016~2021은 `=KET2IMG.FLOAT("...")`) 수식으로 삽입됩니다.
4. 추후 해당 셀을 **더블클릭**하면 Ketcher 2D 가 다시 열리며 화살표 위치·각도·시약 주석이 그대로 복원됩니다.

::: tip `=SMILES2IMG` 에서도 KET JSON 을 자동 감지합니다
한 열에 단일 SMILES 문자열과 Ketcher KET 문자열이 섞여 있더라도 `=SMILES2IMG(A2)` 가 KET JSON 을 자동으로 판별해 반응식 렌더러로 처리합니다.
:::

## 2. 반응식에도 공용 스타일(`[bg]`, `[style]`) 바로 적용

`=KET2IMG` 와 `=KET2IMG.FLOAT` 역시 `=SMILES2IMG` 와 완전히 동일한 스타일 파서를 공유하므로, 다시 그리지 않고도 투고용 흑백이나 다크 모드 반응식으로 즉시 전환됩니다:

| 활용 상황 | 엑셀 수식 | 결과 |
| :--- | :--- | :--- |
| **표준 컬러 (투명 배경)** | `=KET2IMG(A2)` | 원소별 색상 (`O` 빨강, `N` 파랑) + 투명 배경 기본 렌더링 |
| **논문 투고용 ACS 흑백** | `=KET2IMG(A2, "", "bw")` *(단축형 `=KET2IMG(A2, "bw")`)* | 학술지 투고 및 특허 명세서용 순수 흑백 반응식 |
| **다크 슬라이드 / 대시보드용 흰색** | `=KET2IMG(A2, "#0F172A", "white")` | 어두운 배경 위 선명한 흰색 화살표·결합선·텍스트 |
| **구버전 Excel (2016~2021)** | `=KET2IMG.FLOAT(A2, "", "bw")` | 셀 크기 연동 부동 반응식 그림 (`xlMoveAndSize`) |

## 3. 단일 분자 KET 셀에서 분자량 · 물성 직접 계산하기

Ketcher에서 특정 입체 배치나 좌표를 직접 그려 단일 분자 KET JSON 으로 저장한 셀도 일반 SMILES 셀처럼 물성·식별자 수식에 바로 연결할 수 있습니다:

* **화학식 · 분자량 · 정밀질량:** `=KET2FORMULA(A2)`, `=KET2MW(A2)`, `=KET2EXACTMASS(A2)`
* **SMILES · 이름 · InChI · InChIKey:** `=KET2SMILES(A2)`, `=KET2NAME(A2)`, `=KET2INCHI(A2)`, `=KET2INCHIKEY(A2)`
* **LogP · TPSA · 3D 지표:** `=KET2LOGP(A2)`, `=KET2TPSA(A2)`, `=KET2HEAVY(A2)`, `=KET2RINGS(A2)`, `=KET2AROMATIC(A2)`, `=KET2FSP3(A2)`, `=KET.MATCH(A2, "smarts")` *(`SMILES2*` / `SMILES.MATCH` 와 1:1 동등 지원)*

## 관련 가이드

- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — 8개 이미지 함수 공용 `[bg]`, `[style]`, `[transform]` 키워드 전체 정리
- [**SMILES · InChI**](/ko/excel/structures) — 대량 SMILES/InChI 구조식 표 및 더블클릭 편집
- [**Ketcher 2D 편집기 상세 매뉴얼**](/ko/editor/) — 반응 화살표, 원자 매핑, 입체 중심 그리기 단축키 안내
