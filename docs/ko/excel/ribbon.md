# 리본 · 우클릭 메뉴

Excel 상단 **`[SMILES]`** 리본 탭의 모든 명령과 3단계 **`Chemical Structure`** 우클릭 컨텍스트 메뉴 레퍼런스입니다.

<FeatureMedia id="excel-ribbon-tools" />

## 상단 `[SMILES]` 리본 명령

| 그룹 | 명령 | 기능 |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit (SMILES / InChI / KET)` (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>) | 선택한 셀에 대해 Ketcher를 열고 해당 형식 수식을 삽입 |
| | `3D Structure (Mol*)` | Mol\* 3D 뷰어 실행 (셀에 4자리 PDB 코드가 있으면 자동 로드) |
| | `Paste Structure` | 클립보드 또는 선택 셀의 SMILES/InChI/KET를 0.1초 만에 구조식 수식으로 변환 |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG` | 화학 식별자 또는 무손실 벡터 SVG 복사 |
| | `Molecule Info` | 17종 물성 창 열기 또는 TSV · 요약 · 목록 · 마크다운 · JSON 복사 |
| | `Batch Style` | 선택 영역 또는 시트 전체 일괄 스타일 변환 (`Standard Color`, `Publication B&W`, `White Background`, `Dark Mode`) |
| | `SDF File` | 12열 `.sdf` 시트 가져오기 및 V2000 `.sdf` 내보내기 |
| **Diagnostics & About** | `Diagnostics` | EPAM Indigo 파서 오류 로그 확인 (`#VALUE!` 원인 진단) |
| | `About func_SMILES2IMG` | 버전, 서드파티 라이선스 확인 및 업데이트 체크 |

## 우클릭 컨텍스트 메뉴

![Excel 3단계 Chemical Structure 우클릭 메뉴](/excel-context-menu-guide.png)
