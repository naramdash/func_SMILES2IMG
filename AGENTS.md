# AGENTS.md — func_SMILES2IMG Developer & Agent Guidelines

이 문서는 **func_SMILES2IMG** (Excel, Word, PowerPoint) 프로젝트를 개발하거나 유지보수하는 모든 AI 어시스턴트(Agent) 및 개발자가 반드시 준수해야 하는 핵심 아키텍처 규칙과 주의사항을 정의합니다.

---

## 0. 공식 제품명 단일 통일 규칙 (Mandatory Product Naming Rule)

- 본 프로젝트는 Excel, Word, PowerPoint 전 제품을 통틀어 공식 명칭을 **`func_SMILES2IMG`** 하나로 통일합니다.
- **절대 혼용 금지 명칭**: `SMILES for Office`, `SMILES for Microsoft Office Suite`, `SMILES for Excel`, `SMILES for Word`, `SMILES for PowerPoint` 등의 임의 브랜드명을 문서, UI, 설치 스크립트, 릴리스 노트, 빌드 아티팩트에 사용하지 마십시오.
- **적용 범위**:
  1. **제품 UI 및 코드 (`src/`)**: 리본 메뉴(`About func_SMILES2IMG`), `AboutDialog`, 메시지 박스 타이틀(`func_SMILES2IMG`), COM Add-in 등록 `FriendlyName` (`func_SMILES2IMG (Word)`, `func_SMILES2IMG (PowerPoint)`).
  2. **문서 사이트 (`docs/`, `README.md`, `THIRD_PARTY_NOTICES.md`)**: VitePress 사이트 타이틀, 히어로 타이틀, 본문, 이미지 대체 텍스트(`alt`), 메타데이터 전체에서 **`func_SMILES2IMG`**로 표기합니다.
  3. **빌드 아티팩트 및 GitHub Release (`gh release`, `.github/workflows/release.yml`, `scripts/Package-Release.ps1`)**:
     - 원클릭 설치/제거 스크립트: `Install-func_SMILES2IMG.cmd`, `Uninstall-func_SMILES2IMG.cmd`
     - 배포 패키지 아카이브: `func_SMILES2IMG-x64.zip`, `func_SMILES2IMG-x86.zip`
     - GitHub Release 제목 및 본문 헤더: `func_SMILES2IMG vX.Y.Z`

---

## 1. Office Fluent Ribbon UI & `imageMso` 동작 원리 및 아이콘 누락 방지 규칙

### ⚠️ 아이콘이 사라지거나 빈 공간으로 나오는 원인 (Root Cause)

Office Fluent UI (RibbonX XML)에서 아이콘(`imageMso`)을 변경할 때 아이콘이 아예 표시되지 않고 **투명한 빈 사각형(16×16 또는 32×32 공백)**으로 렌더링되는 현상은 Microsoft Office의 의도된 설계 동작입니다:

1. **무음 실패(Silent Failure) 설계**:
   - Office 리본 파서는 `imageMso`에 잘못되거나 존재하지 않는 이름이 지정되어도 **절대 오류 대화상자나 예외를 발생시키지 않고 조용히 투명 처리**합니다.
   - Excel 옵션의 *"추가 기능 사용자 인터페이스 오류 표시"*를 켜도 `imageMso` 누락은 XML 문법 오류가 아니므로 경고창이 뜨지 않습니다.
2. **컨트롤 타입 불일치 (Control Type Mismatch)**:
   - Microsoft Office 컨트롤 ID 목록에는 버튼(`button`) 외에도 갤러리(`gallery`), 콤보박스(`comboBox`), 분할 버튼(`splitButton`) 등이 섞여 있습니다.
   - 예: `FontName`, `ShapesGallery`, `StylesGallery` 등은 드롭다운 콤보박스나 갤러리 ID이므로 `<button imageMso="..." />`의 단독 아이콘 면(face)으로 사용할 수 없어 빈칸이 됩니다.
3. **호스트 애플리케이션 불일치 (Host Application Mismatch)**:
   - Microsoft가 배포하는 ID 목록은 Access, Outlook, Word, Excel, PowerPoint 통합 목록입니다.
   - 예: `Key`(열쇠 아이콘)는 **Microsoft Access 테이블 디자인 전용** 명령어로, Excel/Word/PowerPoint 리본에서는 지원되지 않아 완전히 누락됩니다.
4. **추측된 영단어 (Non-Existent Guessed Names)**:
   - 개발자/에이전트가 영어 단어를 직관적으로 추측하여 넣으면 100% 누락됩니다.
   - 예: `HighContrast`(다크모드용 추측), `DocumentSummary`(요약용 추측), `TableCustom`(마크다운 표 추측), `SourceCodeInsert`(JSON용 추측), `FormatBullet`(목록용 추측), `ShapeFillColorPicker`(배경용 추측), `Tag` 등은 Office 버튼 아이콘으로 존재하지 않습니다.
5. **COM Automation out-of-process 제약**:
   - `Application.CommandBars.GetImageMso`는 현대 Office 365 Click-to-Run 환경에서 out-of-process COM으로 호출 시 Ribbon UI가 초기화되지 않아 `0x8000FFFF (E_UNEXPECTED)` 예외를 발생시킵니다. 따라서 유닛 테스트는 정적 검증과 검증된 카탈로그를 사용해야 합니다.

---

### 🛡️ 에이전트 필수 준수 규칙 (Mandatory Rules)

1. **절대로 `imageMso` 이름을 임의로 추측하지 마십시오.**
   - 사전에 실제 Office(Excel/Word/PowerPoint)에서 정상 렌더링이 확인된 `VerifiedValidOfficeImageMsos` 목록에 있는 ID만 사용하십시오.
2. **모든 리본 변경 시 `dotnet run --project tests`를 반드시 실행하십시오.**
   - `tests/Program.cs`의 `ValidateRibbonXml` 메서드는 Excel, Word, PowerPoint의 리본 및 컨텍스트 메뉴 XML을 파싱하여:
     - `KnownInvalidImageMsos` (블랙리스트)에 포함된 아이콘이 있는지 검사합니다.
     - `VerifiedValidOfficeImageMsos` (화이트리스트)에 등록되지 않은 임의의 아이콘이 있는지 검사하여 즉시 빌드를 중단시킵니다.
3. **새로운 `imageMso`를 추가할 때의 절차**:
   - 먼저 해당 ID가 Excel 2016~365의 표준 `<button>`에서 실제로 표시되는지 확인하십시오.
   - `tests/Program.cs`의 `VerifiedValidOfficeImageMsos`와 본 문서의 목록에 등록하십시오.

---

### 🎨 Batch Style (배경 및 스타일 프리셋) 사각형 아이콘 규격

사용자의 혼란을 방지하기 위해 일괄 스타일(Batch Style) 메뉴의 아이콘은 아래와 같이 **직관적인 사각형/팔레트 표준 아이콘**으로 고정합니다:

| 메뉴 항목 | `imageMso` | 시각적 형상 및 의미 |
| :--- | :--- | :--- |
| **Standard Color** | `ColorMenu` | 표준 색상 팔레트 (컬러 렌더링) |
| **Publication B/W** | `BlackAndWhiteGrayscale` | 반반 흑백/그레이스케일 사각형 (흑백 논문용) |
| **White Background** | `BlackAndWhiteLightGrayscale` | **밝은 흰색 사각형** (흰색 불투명 배경) |
| **Dark Mode** | `BlackAndWhiteBlack` | **순수 검은색 사각형** (다크모드 / 흰색 잉크) |

---

### 📋 프로젝트에서 100% 검증된 `imageMso` 카탈로그 (32종)

| 분류 | `imageMso` | 적용 위치 및 용도 |
| :--- | :--- | :--- |
| **구조 삽입** | `PensGallery` | `btnInsertMolecule` (SMILES 구조 그리기 / 편집) |
| | `TextBoxInsert` | `btnInsertInchiMolecule` (InChI 텍스트 구조 입력) |
| | `SmartArtInsert` | `btnInsertKetMolecule` (KET 캔버스/반응식 다이어그램) |
| | `Insert3DModelFromFile` | `btnInsertMolstar` (Mol* 3D 거대분자 뷰어) |
| | `Paste` | `btnQuickInsert` (클립보드 구조 즉시 붙여넣기) |
| **복사 도구** | `Copy` | `btnCopySmiles`, `btnCopyKet` (표준 텍스트/데이터 복사) |
| | `SignatureLineInsert` | `btnCopyInchi` (공식 표준 인증 식별자 InChI 서명 복사) |
| | `CopyPicture` | `btnCopySvg` (벡터 그래픽 SVG 그림 복사) |
| **분자 정보** | `FileProperties` | `menuRibbonExcelMoleculeInfo` (분자 물성 드롭다운 & 상세창) |
| | `FontSize` | `btnRibbonExcelCopyName` (화합물 명칭 텍스트 복사) |
| | `Subscript` | `btnRibbonExcelCopyFormula` (화학식 $X_2$ 아래첨자 복사) |
| | `Calculator` | `btnRibbonExcelCopyMW` (분자량 계산값 복사) |
| | `FilePermission` | `btnRibbonExcelSubCopyInchiKey` (InChIKey 보안 키 복사) |
| | `FileDocumentInspect` | `btnRibbonExcelCopyAllSummary` (전체 요약 검사/복사) |
| | `TableInsert` | `btnRibbonExcelCopyTsv` (TSV 엑셀 행 복사) |
| | `Bullets` | `btnRibbonExcelCopyMultiLine` (글머리 기호 목록 복사) |
| | `TableOfContentsGallery` | `btnRibbonExcelCopyMarkdown` (마크다운 표 복사) |
| **스타일 변환** | `PictureRecolorGallery` | `menuRibbonExcelBatchStyle` (배치 스타일 드롭다운) |
| | `ColorMenu` | `btnBatchStyleColorTrans` (표준 컬러) |
| | `BlackAndWhiteGrayscale` | `btnBatchStyleBw` (학술 논문용 흑백) |
| | `BlackAndWhiteLightGrayscale` | `btnBatchStyleWhiteBg` (흰색 사각형 - 흰색 배경) |
| | `BlackAndWhiteBlack` | `btnBatchStyleDark` (검은색 사각형 - 다크모드) |
| **SDF 파일** | `DatabaseQueryNew` | `menuRibbonExcelSdf` (화학 DB SDF 메뉴) |
| | `ImportTextFile` | `btnRibbonExcelImportSdf` (SDF 파일 가져오기) |
| | `FileSaveAs` | `btnRibbonExcelExportSdf` (시트 데이터 SDF 내보내기) |
| **진단 & 정보** | `WatchWindow` | `btnRibbonExcelDiagnostics` (시스템 상태 모니터링 안경) |
| | `Info` | `btnRibbonExcelAbout` (버전 및 정보창) |
| **컨텍스트 메뉴** | `DesignMode` | 우클릭 분자 편집 메뉴 |
| | `EquationInsertNew` | 우클릭 수식 삽입 메뉴 |
| | `FunctionWizard` | 우클릭 함수 지원 메뉴 |
| | `GroupTaskProperties` | 우클릭 속성 정보 메뉴 |
| | `PropertySheet` | 우클릭 그리드 시트 메뉴 |

---

## 2. Molecule Info 창 그리드 행별 벡터 아이콘 규칙

`MoleculeInfoDialog`의 `DataGridView`는 외부 이미지 파일에 의존하지 않고 `src/Smiles2Img.Core/Editor/UiIcons.cs`의 **GDI+ 벡터 렌더러**를 통해 20×20 비트맵을 인메모리로 생성하여 사용합니다.

- **DPI 인식**: 시스템 배율(100%, 125%, 150%, 200%)에 맞춰 선 두께와 패딩을 자동 계산합니다.
- **아이콘 매핑**:
  - Substance Name / IUPAC Name: 파란색 문서 & 이름표
  - Molecular Weight: 보라색 저울 / 추
  - Chemical Formula: 오렌지색 원자 핵과 전자 궤도
  - Monoisotopic Mass: 초록색 질량 스펙트럼 피크
  - LogP / TPSA: 파란색 수용성 물방울
  - Lipinski Rule of 5: 빨간색/주황색 알약(캡슐)
  - Canonical SMILES: 연두색 화학 결합 노드
  - InChI / InChIKey: 황금색 열쇠 & 보안 자물쇠

---

## 3. 릴리스 및 배포 절차 (Release Checklist)

버전을 올리거나 미리보기(preview) 릴리스를 발행할 때는 다음 파일들을 **동시에** 수정해야 합니다:

1. `Directory.Build.props`: `<Version>X.Y.Z-preview.N</Version>`
2. `package.json`: `"version": "X.Y.Z-preview.N"`
3. `README.md`: 설치 가이드 및 버전 번호
4. `docs/.vitepress/config.mts`: 최신 버전 뱃지
5. `docs/index.md` & `docs/ko/index.md`
6. `docs/guide/getting-started.md` & `docs/ko/guide/getting-started.md`

### 검증 및 배포 커맨드:
```powershell
# 1. 빌드 및 테스트 통과 필수
dotnet test -c Release
dotnet run --project tests -c Release

# 2. 커밋 및 태그 푸시 -> GitHub Actions 자동 빌드 및 릴리스 생성
git add -A
git commit -m "feat/fix: description"
git tag -a vX.Y.Z-preview.N -m "Release description"
git push origin dev --tags
```
