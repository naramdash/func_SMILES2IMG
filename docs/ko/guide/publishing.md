# 출판 스타일 · SVG 벡터 · 물성 표 내보내기

학술 논문 투고, 특허 명세서 작성, 학회 포스터 및 발표 슬라이드 제작을 마무리할 때 **Excel, Word, PowerPoint 세 프로그램 모두에서 동일하게 사용할 수 있는 3대 출판·내보내기 도구**를 안내합니다.

---

## 1. 🎨 문서 전체 스타일 원클릭 통일 (`Batch Style`)

논문을 투고하거나 발표 슬라이드 테마를 바꿀 때, 문서 안에 들어있는 수십 개의 분자 구조식을 하나하나 다시 그릴 필요가 없습니다. 상단 리본 메뉴의 **`[SMILES] ➔ [Batch Style]`**을 클릭하면 선택한 영역(또는 문서/시트/슬라이드 전체)의 모든 구조식이 1초 만에 통일됩니다!

<BatchStylePreview />

### 프로그램별 적용 범위
* **📗 Excel:** 현재 마우스로 드래그해 선택한 셀 영역의 구조식만 변환하거나, 단일 셀만 선택된 경우 **현재 시트 전체**의 구조식 수식을 일괄 변환합니다.
* **📘 Word:** 드래그해 선택한 문단/표 안의 구조식만 변환하거나, 선택 영역이 없으면 **워드 문서 전체**의 구조식을 일괄 변환합니다.
* **📙 PowerPoint:** 선택한 슬라이드 도형들만 변환하거나, 선택된 도형이 없으면 **프레젠테이션 전체 슬라이드**의 구조식을 일괄 변환합니다.

---

## 2. 🖼️ Illustrator · 포스터용 무손실 벡터 복사 (`Copy SVG`)

일반 화면 캡처(PNG/JPG 비트맵)는 학회 포스터처럼 크게 확대하거나 인쇄할 때 가장자리가 계단처럼 깨집니다.  
**`[SMILES] ➔ [Copy SVG]`** (또는 우클릭 ➔ **`Chemical Structure ➔ Copy Vector Graphic (SVG)`**)를 사용하면 **화질 저하가 전혀 없는 수학적 벡터 그래픽(`SVG`)**으로 클립보드에 복사됩니다.

| 붙여넣을 프로그램 | 활용 방법 및 장점 |
| :--- | :--- |
| **Adobe Illustrator / Inkscape** | <kbd>Ctrl</kbd> + <kbd>V</kbd>로 바로 붙여넣으면 개별 결합선(Path)과 원자 텍스트가 살아있는 벡터 객체로 들어갑니다. 선 굵기나 색상을 자유롭게 재편집할 수 있습니다. |
| **학회 포스터 (PowerPoint / 대판 인쇄)** | A0 크기의 대형 학회 포스터로 수십 배 확대해 인쇄해도 결합선과 글자가 칼날처럼 선명하게 유지됩니다. |
| **웹 에디터 / 코드 편집기 (VS Code 등)** | 클립보드에 그래픽 스트림(`image/svg+xml`)과 원본 `<svg>...</svg>` XML 텍스트가 동시에 복사되므로, 텍스트 에디터에 붙여넣으면 깔끔한 SVG 코드를 바로 얻을 수 있습니다. |

---

## 3. 🔬 분자 정보 상세창 (`Molecule Info`) & 5가지 서식 내보내기

구조식 셀이나 그림을 **우클릭 ➔ `Chemical Structure ➔ Copy Molecule Info`** (또는 상단 리본의 **`[Molecule Info]`**)를 누르면 외부 검색 없이 분자의 모든 물리화학적 특성을 확인하고 원하는 서식으로 즉시 복사할 수 있습니다.

### 한눈에 확인 가능한 10가지 핵심 물성
1. **Substance Name (물질명):** IUPAC / PubChem 일반명 (`Aspirin`)
2. **Chemical Formula (화학식):** 표준 Hill 표기법 (`C9H8O4`)
3. **Molecular Weight (분자량):** 평균 분자량 (`180.16 g/mol`)
4. **Monoisotopic Mass (정밀 질량):** 질량분석(HRMS)용 단일동위원소 질량 (`180.0423 Da`)
5. **LogP & TPSA:** 옥탄올/물 분배계수(`1.31`) 및 위상 극성 표면적(`63.60 Å²`)
6. **Lipinski Rule of 5:** 경구 투여 약물성 통과 여부 및 HBD(`1`), HBA(`4`), 회전 결합 수(`3`)
7. **표준 식별자 3종:** Canonical SMILES, IUPAC InChI, 27자리 InChIKey

### 📋 붙여넣을 곳에 맞춰 골라 쓰는 5가지 일괄 복사 서식

우클릭 메뉴나 상세창 하단의 **[내보내기 (Export)]** 버튼에서 아래 5가지 서식 중 하나를 선택해 복사할 수 있습니다:

#### ① 엑셀 가로 행 / 워드 표 칸 자동 분할 (`Copy as Excel Row — TSV Columns`)
탭 구분(TSV) 및 HTML 표 형식으로 복사됩니다. **엑셀 셀이나 워드/파워포인트 표(Table)에 <kbd>Ctrl</kbd> + <kbd>V</kbd>로 붙여넣으면 각 항목이 옆 칸(열)으로 자동 분할**되어 표가 한 번에 채워집니다!
```text
Aspirin	C9H8O4	180.16 g/mol	180.0423 Da	CC(=O)Oc1ccccc1C(=O)O	BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ② 파워포인트 슬라이드 글머리 기호용 (`Copy as Multi-Line List`)
슬라이드 텍스트 상자에 붙여넣으면 즉시 깔끔한 불릿(Bullet) 요약 목록이 됩니다:
```text
• Name: Aspirin
• Formula: C9H8O4
• Molecular Weight: 180.16 g/mol
• Exact Mass: 180.0423 Da
• LogP / TPSA: 1.31 / 63.60 Å²
• Lipinski Ro5: PASS (HBD: 1, HBA: 4, RotBonds: 3)
• SMILES: CC(=O)Oc1ccccc1C(=O)O
• InChIKey: BSYNRYMUTXBXSQ-UHFFFAOYSA-N
```

#### ③ 노션(Notion) · GitHub · 전자연구노트(ELN)용 (`Copy as Markdown Table`)
```markdown
| Property | Value |
| :--- | :--- |
| **Substance Name** | Aspirin |
| **Chemical Formula** | C9H8O4 |
| **Molecular Weight** | 180.16 g/mol |
| **Exact Mass** | 180.0423 Da |
| **LogP / TPSA** | 1.31 / 63.60 Å² |
| **Lipinski Ro5** | PASS (HBD: 1, HBA: 4, RotBonds: 3) |
```

#### ④ 논문 본문 · 이메일 각주용 한 줄 요약 (`Copy All Summary`)
```text
Aspirin | Formula: C9H8O4 | MW: 180.16 g/mol | Exact Mass: 180.0423 Da | SMILES: CC(=O)Oc1ccccc1C(=O)O
```

#### ⑤ 파이썬(Python) · 데이터 분석 연동용 (`Export as JSON`)
상세 정보 창(`Detailed Molecule Info Window...`)에서 JSON 내보내기를 선택하면 모든 물리화학 지표가 구조화된 JSON 포맷으로 복사됩니다.
