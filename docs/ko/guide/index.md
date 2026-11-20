# 소개 (Introduction)

`func_SMILES2IMG`는 **Microsoft Excel, Word, PowerPoint**에서 화학 2D 구조식과 3D 단백질 모델을 바로 다루는 무료·오픈소스 추가 기능입니다.

엑셀 셀에 화합물 이름, CAS 번호, SMILES를 입력하면 수백 행의 2D 구조식과 분자 물성이 자동으로 그려지며, Word와 PowerPoint에서도 내장 **Ketcher 2D** 편집기와 **Mol\* 3D** 단백질 뷰어로 더블클릭 제자리 편집을 할 수 있습니다.

## Microsoft Office 앱별 핵심 기능

| Office 앱 | 주요 기능 | 가이드 |
| :--- | :--- | :--- |
| **Microsoft Excel** | 셀 안 구조식 수식(`=SMILES2IMG`, `=NAME2IMG`), 40개 물성·Lipinski·SAR 함수, `.sdf` 입출력, 더블클릭 2D 편집 | [Excel 레시피 북](/ko/excel/) |
| **Microsoft Word** | 논문 본문 및 화합물 표 인라인 구조식, 열 너비 유지 더블클릭 수정, 원클릭 ACS 흑백 변환, 벡터 SVG 복사 | [Word 튜토리얼](/ko/word/) |
| **Microsoft PowerPoint** | 투명 배경 다크 모드 흰색 선화, 텍스트 상자 0.1초 변환, 발표 중 360° 회전 3D 단백질 모델(`.glb`) | [PowerPoint 갤러리](/ko/powerpoint/) |
| **공통 에디터** | 3개 앱 모두에서 공통 작동하는 오프라인 **Ketcher 2D** 편집기(`Ctrl+Shift+K`) 및 **Mol\* 3D** 단백질 뷰어 | [Editor 가이드](/ko/editor/) |

![Microsoft Excel에서 실행 중인 func_SMILES2IMG](/excel-showcase.png)

## 기존 도구와 비교

| 비교 항목 | **func_SMILES2IMG** | **ChemDraw / JChem for Office** | **Chem4Word** |
| :--- | :--- | :--- | :--- |
| **라이선스 및 비용** | **무료 · 오픈소스 (MIT)** | 유료 상용 구독 라이선스 필요 | 무료 오픈소스 |
| **지원 오피스 앱** | **Excel + Word + PowerPoint 통합** | Excel + Word + PowerPoint | Word 전용 |
| **엑셀 구조식 방식** | **네이티브 인셀 이미지(`=SMILES2IMG`)** 및 `.FLOAT` 도형 (별도 시트 변환 불필요) | 전용 시트 변환(Convert Worksheet) 또는 전용 객체 필요 | 미지원 |
| **3D 단백질 뷰어** | **Mol\* 내장 (360° 회전 `.glb` 모델 & 2D 캡처)** | 별도 3D 프로그램 필요 | 미지원 |
| **설치 권한** | **관리자 권한 불필요** (10초 유저 영역 설치) | 관리자(Admin) 설치 및 라이선스 인증 필요 | 설치 마법사 실행 |

## 로컬 우선(Local-First) 보안 & 문서 공유 {#privacy}

### 로컬 처리 vs. 선택적 네트워크 사용
핵심 2D 구조식 렌더링(**EPAM Indigo**), 2D 스케처(**EPAM Ketcher**), 3D 뷰어 런타임(**Mol\***), 모든 물성·Lipinski·SAR 계산, `.sdf` 입출력, Batch Style, SVG 복사는 **내 PC에서 로컬로 실행**되며 구조식을 외부 서버로 전송하지 않습니다.

아래 선택 기능을 직접 호출할 때만 인터넷에 접속합니다:
1. **PubChem 물질명 / CAS 조회:** 내장 사전에 없는 화합물 이름·CAS 번호를 조회할 때(`=NAME2IMG`, `=NAME2SMILES`, `=SMILES2NAME`, `Molecule Info` 물질명) PubChem에 1회 조회하고 결과를 디스크에 로컬 캐시합니다.
2. **Mol\* PDB 코드 다운로드:** 4자리 PDB 코드(`1CRN`, `6LU7` 등)로 공개 3D 구조를 불러올 때 RCSB PDB / PDBe에서 다운로드합니다(로컬 `.pdb`/`.cif` 파일 열기는 오프라인 작동).
3. **업데이트 확인 및 온라인 설치:** `About` 창에서 **`[Check for Updates]`** 버튼을 누르거나 `install.ps1`을 실행할 때 GitHub Releases에 접속합니다.

### 추가 기능이 없는 PC / Mac에서 문서 열기
* **Word (`.docx`) & PowerPoint (`.pptx`):** 2D 구조식은 표준 고해상도 PNG 그림으로, 3D 단백질은 표준 `.glb` 3D 모델로 문서에 저장됩니다. 추가 기능이 설치되지 않은 Windows, **macOS**, 아이패드, 웹 오피스에서도 깨짐 없이 열람·인쇄·PDF 저장이 가능합니다.
* **Excel (`.xlsx`):** 이미 렌더링된 이미지 캐시와 계산 결과는 다른 PC에서도 표시됩니다. 재계산(<kbd>F9</kbd>)이 발생할 수 있는 외부 배포 시에는 수식 열을 **값으로 붙여넣기**하거나 **PDF로 저장**해 전달하는 것을 권장합니다.

## 시스템 요구사항

| 항목 | 지원 사양 | 비고 |
| :--- | :--- | :--- |
| **운영체제** | Windows 10 / Windows 11 | 32비트 및 64비트 모두 지원 |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | 데스크톱 설치 버전 |
| **필수 런타임** | .NET Framework 4.8 & Edge WebView2 | Windows 10/11 기본 내장 |

## 다음 단계

- [**설치 (10초 원클릭 셋업)**](/ko/guide/install) — 관리자 권한 없이 PowerShell 한 줄로 Excel · Word · PowerPoint 전체에 자동 설치합니다.
- [**시작하기 (1분 첫 분자 그리기)**](/ko/guide/getting-started) — Excel, Word, PowerPoint에서 첫 구조식을 만들어 보고 8개 탭 샘플 워크북을 받아보세요.
