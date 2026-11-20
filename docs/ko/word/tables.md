# 화합물 표 · 물성 복사

Word 화합물 요약 표에 구조식을 배치하고, 화학식·분자량·LogP·Lipinski 판정을 수동 계산 없이 바로 복사해 채워 넣습니다.

<FeatureMedia id="word-ch3-compound-table" />

## 사용 순서

1. Word 표 셀 안에 커서를 두고 **`[Insert / Edit Structure]`** 또는 **`[Paste Structure]`** 버튼으로 각 화합물을 삽입합니다.
2. 삽입된 구조식을 선택한 뒤 **`[SMILES]` ➔ `[Molecule Info]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Copy Molecule Info`**)를 클릭합니다.
3. 표 형태에 맞는 복사 서식을 선택합니다:
   - **`Copy All Summary (Text)`:** 머리글(Header) + 데이터 2줄 표로 복사 (새 표 생성에 적합)
   - **`Copy as Excel Row (TSV Columns)`:** 헤더 없이 17개 탭 구분 데이터만 복사 (기존 표 행 채우기에 적합)
   - **`Detailed Molecule Info Window...`:** 2D 미리보기와 함께 17종 물성을 확인하고 개별 행 더블클릭 복사 또는 Markdown/JSON 복사

::: tip 재편집 시 표 열 너비 고정 팁
표 우클릭 ➔ **자동 맞춤** ➔ **열 너비 고정**을 설정해 두면 나중에 구조식을 더블클릭해 수정해도 옆 텍스트 열이 밀리지 않습니다.
:::
