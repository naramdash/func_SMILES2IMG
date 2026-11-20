# 시작하기 (Getting Started)

**`func_SMILES2IMG`** 설치 후 상단에 **`[SMILES]`** 리본 탭이 보인다면, 1분 안에 Excel · Word · PowerPoint에서 첫 분자 구조식을 그려보세요.

## 1. Microsoft Excel — 첫 수식 입력 (`=NAME2IMG` / `=SMILES2IMG`)

아무 빈 셀에나 아래 수식을 붙여넣고 <kbd>Enter</kbd>를 눌러보세요(**Excel 2016~2021**에서는 뒤에 `.FLOAT`를 붙입니다: `=NAME2IMG.FLOAT("Aspirin")`):

```text
=NAME2IMG("Aspirin")
```
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```

행 높이(예: `62~80 pt`)와 열 너비를 넉넉하게 키우면 가로세로 비율이 유지된 채 선명하게 확대됩니다. 구조식 셀을 **더블클릭**하면 내장 Ketcher 2D 편집기가 열려 즉시 수정할 수 있습니다.

<FeatureMedia id="excel-overview" />

::: tip 자체 완결형 공식 Playbook 다운로드
웹 문서 없이 Excel · Word · PowerPoint 안에서 바로 실습해 보고 싶다면 [예제 파일 (Playbooks)](/ko/guide/playbooks) 페이지에서 3종 공식 Playbook(`.xlsx`, `.docx`, `.pptx`)을 내려받으세요.
:::

## 2. Microsoft Word — 본문 & 화합물 표 삽입

1. 본문이나 화합물 요약 표 안의 원하는 위치에 커서를 둡니다.
2. 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (또는 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>)로 Ketcher 편집기를 엽니다.
3. 구조식이나 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 인라인 그림으로 삽입됩니다. 나중에 그림을 **더블클릭**해 수정해도 표 열 너비가 깨지지 않습니다.

<FeatureMedia id="word-overview" />

## 3. Microsoft PowerPoint — 다크 모드 & 3D 단백질 슬라이드

어두운 발표 템플릿에 흰색 테두리 없이 투명 배경 구조식을 배치하고, SMILES·화합물 이름이 적힌 텍스트 상자를 **`[Paste Structure]`** 버튼으로 0.1초 만에 변환하거나 **`[3D Structure (Mol*)]`** 버튼으로 360° 회전하는 3D 단백질을 삽입합니다.

<FeatureMedia id="powerpoint-overview" />

## 4. 앱별 전용 가이드로 계속하기

- [**Microsoft Excel — 실무 레시피 북**](/ko/excel/) — 구조식 표, 이름/CAS 변환, Lipinski 약물성, SAR 스크리닝 레시피와 40개 함수 사전
- [**Microsoft Word — 6장 논문 튜토리얼**](/ko/word/) — 실제 연구 논문 원고 1편을 서론 삽입부터 투고용 ACS 흑백 통일까지 완성해 나갑니다
- [**Microsoft PowerPoint — 슬라이드 갤러리**](/ko/powerpoint/) — 다크 모드 투명 구조식, 반응 메커니즘, 360° 회전 3D 단백질 슬라이드를 보고 따라 합니다
- [**Editor — Ketcher 2D · Mol\* 3D · 출력**](/ko/editor/) — 2D 구조·반응식 편집기, 3D 단백질 뷰어, ACS 흑백 일괄 변환 및 벡터 SVG 복사 가이드
