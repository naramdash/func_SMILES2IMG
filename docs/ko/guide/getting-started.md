# 시작하기 및 설치 안내

**SMILES for Microsoft Office Suite**는 Microsoft Excel, Microsoft Word, Microsoft PowerPoint에서 2D 화학 구조식 및 3D 생체 거대분자(Mol*)를 고해상도로 렌더링하고, 분자 특성을 계산하며, 시각적으로 편집할 수 있는 화학 연구용 통합 추가 기능(Add-in)입니다.

---

## ⚡ 빠른 설치 (권장)

**PowerShell**을 열고 아래 명령어를 붙여넣은 뒤 Enter 키를 누릅니다 (관리자 권한 없이 일반 사용자 권한으로 즉시 설치 가능):

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 `v1.7.0` 주요 신기능 안내 (정식 안정판)
* **🔬 인셀 SAR 부분구조 & 약물성(Lipinski) 스크리닝:** `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2LIPINSKI`, `=SMILES2LOGP`, `=SMILES2TPSA`로 외부 툴 없이 0ms 실시간 분석
* **🏷️ 화합물명/CAS 번호 역변환:** `=NAME2IMG("Aspirin")`, `=NAME2SMILES("50-78-2")`로 화합물 일반명과 CAS 번호에서 직접 구조식 렌더링
* **🎨 문서 전체 일괄 스타일 변환 (`Batch Style`):** 시트, 슬라이드, 문서 내 모든 구조식을 클릭 한 번으로 **표준 컬러**, **ACS 논문 흑백**, **흰색 배경**, **다크 모드**로 일괄 통일
* **🖼️ 무손실 벡터 SVG 복사 (`Copy SVG`):** Illustrator, Inkscape, ChemDraw에 벡터로 바로 붙여넣는 네이티브 SVG 클립보드 복사
* **📂 화학 데이터베이스 SDF 입출력:** `.sdf` 파일을 읽어 라이브 구조식과 물성 수식이 포함된 엑셀 표로 자동 생성하거나, 시트 데이터를 `.sdf`로 내보내기
* **📋 개편된 분자 정보 창 (Molecule Info):** 물성 행별 고해상도 GDI+ 벡터 케미컬 아이콘, DataGridView 표 레이아웃, 다중 서식 내보내기 Split Button (엑셀 TSV 행, 줄바꿈 목록, 마크다운 표, JSON)
* **⚡ 초고속 재실행 (<0.05초):** WebView2 웜 인스턴스 풀링으로 Ketcher 2D 에디터와 Mol* 3D 뷰어 즉각 오픈
:::

### 설치 스크립트가 자동으로 수행하는 작업
1. **Office 환경 자동 감지:** 설치된 Office가 32비트(`x86`)인지 64비트(`x64`)인지 확인하고 Excel, Word, PowerPoint 설치 여부를 판별합니다.
2. **최신 바이너리 다운로드:** GitHub Releases 최신 통합 배포 패키지를 다운로드합니다.
3. **안전 디렉터리 배치:** `%APPDATA%\Microsoft\AddIns\Smiles2Img` 경로에 파일을 안전하게 복사합니다.
4. **보안 차단 해제:** Windows 다운로드 차단 플래그(`Zone.Identifier`)를 `Unblock-File`로 일괄 자동 해제합니다.
5. **추가 기능 자동 등록:**
   - Excel XLL: `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN` 레지스트리에 자동 등록
   - Word COM Add-in: `HKCU\Software\Microsoft\Office\Word\Addins\Smiles2Img.WordAddin` 등록
   - PowerPoint COM Add-in: `HKCU\Software\Microsoft\Office\PowerPoint\Addins\Smiles2Img.PowerPointAddin` 등록

---

## 📊 실행 요구사항

| 구성 요소 | 요구 버전 | 비고 |
| :--- | :--- | :--- |
| **운영체제** | Windows 10 / Windows 11 | 32비트 및 64비트 환경 완벽 지원 |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | 데스크톱 설치 버전 (간편 실행 및 MSI) |
| **.NET Framework** | .NET Framework 4.8 이상 | Windows 10 (1903+) 및 Windows 11 기본 내장 |
| **Microsoft Edge WebView2** | Evergreen Runtime | 오프라인 Ketcher 2D 에디터 및 Mol* 3D 뷰어 구동용 (최신 Windows 기본 내장) |

---

## 🛡️ Windows 11 스마트 앱 컨트롤 (Smart App Control) 해결법

Windows 11에서 **스마트 앱 컨트롤(SAC)**이 켜져 있는 경우 *"파일 형식과 확장명이 일치하지 않습니다"* 경고가 발생하거나 엑셀 수식이 `#NAME?` 오류로 표시될 수 있습니다.

::: tip 안심하고 진행하세요
스마트 앱 컨트롤을 끄더라도 **Microsoft Defender 백신(실시간 바이러스 감시 및 클라우드 보호)은 100% 정상 작동**합니다.
:::

### 10초 만에 해결하기
1. Windows **설정** (`Win + I`) ➔ **개인 정보 및 보안** ➔ **Windows 보안**으로 이동합니다.
2. **앱 및 브라우저 컨트롤** ➔ **스마트 앱 컨트롤 설정**을 클릭합니다.
3. 설정을 **'꺼짐(Off)'**으로 변경합니다.
4. Excel, Word, PowerPoint를 다시 시작하면 추가 기능이 즉시 정상 작동합니다.

---

## 🖐️ 수동 설치 방법 (기업망 / 오프라인 환경)

보안 정책으로 온라인 PowerShell 스크립트 실행이 차단되는 환경:

1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases)에서 최신 `Smiles2Img-Office-x64.zip` (또는 `Smiles2Img-Office-x86.zip`) 및 `install.ps1`을 다운로드합니다.
2. 다운로드한 `.zip` 파일을 우클릭 ➔ **속성** ➔ 하단 **[차단 해제(Unblock)]** 체크 후 **[확인]**을 누릅니다.
3. 압축을 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 경로에 해제합니다.
4. 해당 폴더에서 PowerShell을 열고 `install.ps1`을 로컬 실행합니다:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

---

## 🔄 버전 확인 및 업그레이드 방법

* **인앱 업데이트 확인:** Excel, Word, PowerPoint 상단 **`[SMILES]`** 리본 탭에서 **`[About SMILES]`** 버튼을 누른 뒤 **`[Check for Updates]`**를 클릭하면 GitHub 최신 릴리즈 버전을 즉시 확인하고 업그레이드 명령어를 복사할 수 있습니다.
* **명령어 업그레이드:** 언제든 빠른 설치 명령어(`irm ... | iex`)를 다시 실행하면 최신 버전으로 덮어쓰기 업그레이드됩니다.

---

## 🗑️ 제거 방법 (Uninstall)

모든 추가 기능 등록과 파일을 완전히 삭제하려면:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
```
또는 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 폴더에서 `uninstall.ps1`을 실행하세요.
