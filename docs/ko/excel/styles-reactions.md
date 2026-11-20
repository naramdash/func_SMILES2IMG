# 스타일 · 하이라이트 · 반응식

2번째(`[bg]`)·3번째(`[style]`) 인자로 셀별 렌더링 스타일을 지정하거나, `=KET2IMG`로 다단계 합성 반응식을 셀 안에 그대로 저장합니다.

## 1. 수식 스타일, 부분구조 하이라이트 & 원자 번호

| 원하는 스타일 | 수식 문법 | 주요 용도 |
| :--- | :--- | :--- |
| **표준 컬러 (투명 배경)** | `=SMILES2IMG(A2)` | IUPAC 원소 색상 (`O` 빨강, `N` 파랑) |
| **논문용 ACS 흑백** | `=SMILES2IMG(A2, "", "bw")` | 학술지 투고 및 특허 명세서용 흑백 선화 |
| **부분구조 컬러 하이라이트** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | 특정 SMARTS 골격만 컬러로 강조 |
| **다크 모드 (흰색 선화)** | `=SMILES2IMG(A2, "#1E293B", "white")` | 어두운 셀 배경 위 선명한 흰색 결합선 |
| **커스텀 선 색상** | `=SMILES2IMG(A2, "", "ink=#1e40af")` | 보고서·테마 브랜드 색상 지정 |
| **NMR 원자 번호 & 입체 표기** | `=SMILES2IMG(A2, "", "num stereo")` | 1번부터 시작하는 원자 번호 및 `R/S`, `E/Z` 표기 |

<FeatureMedia id="excel-style-matrix" />

옵션을 클릭해 미리보면서 수식을 만들려면 [수식 생성기](/ko/excel/formula-builder)를 열어보세요.

## 2. 다단계 합성 반응식 저장 (`=KET2IMG` + Ketcher)

반응 화살표(`➔`), `+`, 촉매 조건이 포함된 합성 반응식을 편집 가능 상태를 유지한 채 엑셀 셀에 삽입합니다:

```text
=KET2IMG(A2)
```

<FeatureMedia id="excel-reactions-ket" />

1. 빈 셀을 선택하고 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>(또는 **`[SMILES]` ➔ `[Insert / Edit (KET)]`**)를 누릅니다.
2. 오프라인 **Ketcher 2D 편집기**에서 반응물, 화살표, 촉매, 생성물을 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 `=KET2IMG(...)`로 삽입됩니다. 추후 셀을 **더블클릭**하면 `0.05초` 이내에 원본 반응식 그대로 다시 열립니다.
