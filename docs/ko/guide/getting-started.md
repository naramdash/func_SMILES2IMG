# 시작하기 및 설치 안내

**SMILES for Microsoft Office Suite**는 Microsoft Excel, Microsoft Word, Microsoft PowerPoint에서 2D 화학 구조식 및 3D 생체 거대분자(Mol*)를 고해상도로 렌더링하고, 분자 특성을 계산하며, 시각적으로 편집할 수 있는 화학 연구용 통합 추가 기능(Add-in)입니다.

---

## ⚡ 빠른 설치 (권장)

**PowerShell**을 열고 아래 명령어를 붙여넣은 뒤 Enter 키를 누릅니다 (관리자 권한 없이 일반 사용자 권한으로 즉시 설치 가능):

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 프리뷰 릴리즈 (`v1.7.0-preview.1`) 설치
부분구조 검색/유사도(`=SMILES.MATCH`, `=SMILES.SIMILARITY`), 약물성 지표(LogP/TPSA/Lipinski), 이름/CAS 번호 역변환(`=NAME2IMG`), 벡터 SVG 복사, 일괄 스타일 변환(`Batch Style`), SDF 파일 입출력(`SDF File`)이 포함된 최신 프리뷰 버전을 설치하려면 아래와 같이 실행하세요:
```powershell
$v='v1.7.0-preview.1'; irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```
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

1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases)에서 최신 `Smiles2Img-vX.X.X.zip`을 다운로드합니다.
2. 다운로드한 `.zip` 파일을 우클릭 ➔ **속성** ➔ 하단 **[차단 해제(Unblock)]** 체크 후 **[확인]**을 누릅니다.
3. 압축을 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 경로에 해제합니다.
4. 해당 폴더에서 PowerShell을 열고 `install.ps1`을 로컬 실행합니다:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

---

## 🗑️ 제거 방법 (Uninstall)

모든 추가 기능 등록과 파일을 완전히 삭제하려면:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
```
또는 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 폴더에서 `uninstall.ps1`을 실행하세요.
