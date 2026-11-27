# 소개

`func_SMILES2IMG`는 **Microsoft Excel, Word, PowerPoint** 안에서 편집 가능한 2D 화학 구조식, 분자 물성 계산, 3D 단백질 모델을 바로 다루는 무료·오픈소스(MIT) 추가 기능입니다.

외부 드로잉 툴에서 정적 이미지를 매번 캡처해 붙여넣는 대신, Excel에서는 수식 한 줄로 수백 행의 구조식과 물성을 자동 생성하고 Word·PowerPoint에서는 더블클릭으로 구조식을 제자리에서 수정할 수 있습니다. 핵심 렌더링과 계산은 모두 **내 PC에서 로컬로 실행**됩니다.

## Microsoft Office 앱별 활용

| Office 앱 | 주요 활용 | 가이드 |
| :--- | :--- | :--- |
| <AppBadge app="excel" /> | 화합물명·CAS·SMILES·InChI를 셀 안 2D 구조식(`=SMILES2IMG`, `=NAME2IMG`)으로 자동 렌더링하고, 49개 물성·3D·변환·SAR 함수 계산과 `.sdf` 가져오기를 한 번에 처리합니다. | <span class="s2i-nowrap">[개요 보기 →](/ko/excel/)</span> |
| <AppBadge app="word" /> | 논문 본문과 화합물 SAR 표에 인라인 구조식을 배치하고, `Molecule Info`로 물성을 정확히 복사하며, ACS 학술지 흑백 일괄 변환과 벡터 SVG 복사를 지원합니다. | <span class="s2i-nowrap">[개요 보기 →](/ko/word/)</span> |
| <AppBadge app="powerpoint" /> | 슬라이드 텍스트 상자를 투명 다크 모드 구조식으로 변환하고, 발표 중 360° 회전하는 3D 단백질(`.glb`) 삽입과 덱 전체 `Batch Style` 전환을 지원합니다. | <span class="s2i-nowrap">[개요 보기 →](/ko/powerpoint/)</span> |

세 앱 모두 오프라인 **[Ketcher 2D](/ko/editor/)** 편집기와 **[Mol\* 3D](/ko/editor/molstar)** 단백질 뷰어가 내장되어 있어, 삽입된 구조식이나 반응식을 더블클릭하면 레이아웃을 깨뜨리지 않고 즉시 수정할 수 있습니다.

![Microsoft Excel에서 실행 중인 func_SMILES2IMG](/excel-showcase.png)

![Microsoft Word에서 실행 중인 func_SMILES2IMG](/word-showcase.png)

![Microsoft PowerPoint에서 실행 중인 func_SMILES2IMG](/powerpoint-showcase.png)

## 기존 도구와 비교

| 비교 항목 | **func_SMILES2IMG** | **ChemDraw / JChem for Office** | **Chem4Word** |
| :--- | :--- | :--- | :--- |
| **라이선스 및 비용** | **무료 · 오픈소스 (MIT)** | 유료 상용 구독 라이선스 필요 | 무료 오픈소스 |
| **지원 오피스 앱** | **Excel + Word + PowerPoint 통합** | Excel + Word + PowerPoint | Word 전용 |
| **엑셀 구조식 방식** | **네이티브 인셀 이미지(`=SMILES2IMG`)** 및 `.FLOAT` 도형 (별도 시트 변환 불필요) | 전용 시트 변환(Convert Worksheet) 또는 전용 객체 필요 | 미지원 |
| **3D 단백질 뷰어** | **Mol\* 내장 (360° 회전 `.glb` 모델 & 2D 캡처)** | 별도 3D 프로그램 필요 | 미지원 |
| **설치 권한** | **관리자 권한 불필요** (유저 영역 설치) | 관리자(Admin) 설치 및 라이선스 인증 필요 | 설치 마법사 실행 |

## 로컬 우선(Local-First) 보안 & 문서 공유 {#privacy}

### 로컬 처리 vs. 선택적 네트워크 사용
핵심 2D 구조식 렌더링(**EPAM Indigo**), 2D 스케처(**EPAM Ketcher**), 3D 뷰어 런타임(**Mol\***), 모든 물성·3D·SAR 계산, `.sdf` 가져오기, Batch Style, SVG 복사는 **내 PC에서 로컬로 실행**되며 구조식을 외부 서버로 전송하지 않습니다.

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

- [**설치 (PowerShell 자동 설치)**](/ko/guide/install) — 관리자 권한 없이 PowerShell 한 줄로 Excel · Word · PowerPoint 전체에 자동 설치합니다.
- [**시작하기 (첫 분자 그리기)**](/ko/guide/getting-started) — Excel, Word, PowerPoint에서 첫 구조식을 만들어 보세요.
- [**예제 파일 (Playbooks)**](/ko/guide/playbooks) — Excel(10개 탭) · Word(6페이지) · PowerPoint(8슬라이드) 자체 완결형 공식 Playbook을 받아보세요.
