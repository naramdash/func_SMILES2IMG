# 설치 및 1분 첫걸음

**func_SMILES2IMG (SMILES for Office)**는 Microsoft Excel, Word, PowerPoint에서 화학 구조식과 반응식을 직접 그리고, 수식으로 분자량·약물성을 계산하며, 3D 단백질을 시각화할 수 있는 무료 추가 기능입니다.

---

## Step 1. 10초 원클릭 설치

1. 실행 중인 **Excel, Word, PowerPoint를 모두 닫습니다.**
2. Windows 시작 메뉴에서 **PowerShell**을 검색해 열고, 아래 명령어를 복사해 붙여넣은 뒤 <kbd>Enter</kbd>를 누릅니다 (관리자 권한 필요 없음):

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

3. `"Installation complete!"` 메시지가 나타나면 설치가 완료된 것입니다!

---

## Step 2. 1분 만에 첫 분자 띄워보기

설치가 끝났다면 오피스를 켜고 아래 3가지를 바로 따라 해보세요.

### ① 엑셀에서 수식으로 구조식 넣기
엑셀을 열고 아무 빈 셀에나 아래 수식을 입력한 뒤 <kbd>Enter</kbd>를 눌러보세요 (행 높이를 조금 넓히면 더 크게 보입니다):

```text
=NAME2IMG("Aspirin")
```
또는 SMILES 문자열이 있다면:
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```
> 💡 **구버전 엑셀(Excel 2016 / 2019 / 2021)을 쓰시나요?**  
> 수식 이름 뒤에 `.FLOAT`를 붙여 `=NAME2IMG.FLOAT("Aspirin")` 또는 `=SMILES2IMG.FLOAT(A2)`로 입력하면 셀 위치에 맞춰 움직이는 그림으로 똑같이 생성됩니다!

### ② 마우스로 직접 구조식·반응식 그리기 (`Ctrl + Shift + K`)
Excel, Word, PowerPoint 어디서든:
1. 상단 리본 메뉴에서 **`[SMILES]`** 탭을 클릭하고 **`[구조식 삽입 (Insert Structure)]`** 버튼을 누르거나, 키보드에서 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누릅니다.
2. 내장된 **Ketcher 2D 화학 편집기**가 열리면 원하는 분자나 반응식을 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (또는 우측 하단 **Apply** 버튼)를 누르면 문서나 셀에 고화질 구조식이 즉시 들어갑니다!

### ③ 더블클릭으로 수정 & 우클릭으로 분자량 복사하기
* **제자리 수정:** 문서나 셀에 들어간 분자 그림을 **마우스로 더블클릭**하면 언제든 편집기가 다시 열려 바로 고칠 수 있습니다.
* **물성 복사:** 분자 그림이나 셀을 **우클릭 ➔ `Chemical Structure` ➔ `Copy Molecule Info`** 메뉴로 선택하면 분자량(`180.16 g/mol`), 화학식(`C9H8O4`), InChIKey, 또는 표 전체를 클립보드로 즉시 복사할 수 있습니다.

---

## Step 3. 다음 단계로 넘어가기

내가 주로 사용하는 프로그램의 가이드를 확인해 보세요:

* 📗 **[Microsoft Excel 가이드](/ko/guide/excel):** 대량 화합물 표 만들기, 분자량·약물성(Lipinski) 자동 계산, 부분구조 하이라이트, `.sdf` 파일 불러오기
* 📘 **[Microsoft Word 가이드](/ko/guide/word):** 논문 본문 및 표 안에 구조식 넣기, 학술지 투고용 흑백(ACS 스타일) 일괄 변환
* 📙 **[Microsoft PowerPoint 가이드](/ko/guide/powerpoint):** 다크 모드 발표 슬라이드용 투명 배경 구조식, 마우스로 360° 돌려보는 3D 단백질 삽입
* 📐 **[엑셀 함수 치트시트](/ko/reference/formula):** 25종 이상의 화학 수식과 렌더링 옵션 한눈에 찾기

---

## 🛠️ 설치 후 실행이 안 될 때 (10초 해결법)

### Windows 11에서 경고가 뜨거나 엑셀 수식이 `#NAME?`으로 나오는 경우
Windows 11의 **스마트 앱 컨트롤**(Smart App Control)이 켜져 있으면 오픈소스 추가 기능 파일을 차단할 수 있습니다.

::: tip 안심하고 진행하세요
스마트 앱 컨트롤을 끄더라도 **Microsoft Defender 백신**(실시간 바이러스 감시 및 클라우드 보호)은 100% 그대로 정상 작동합니다.
:::

1. Windows **설정** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **개인 정보 및 보안** ➔ **Windows 보안**으로 이동합니다.
2. **앱 및 브라우저 컨트롤** ➔ **스마트 앱 컨트롤 설정**을 클릭합니다.
3. 설정을 **'꺼짐(Off)'** 상태로 변경한 뒤 Excel / Word / PowerPoint를 다시 시작하면 즉시 정상 작동합니다.

---

## 🔄 업데이트 및 제거 방법

* **최신 버전 확인 및 업그레이드:**
  - 오피스 상단 **`[SMILES]`** 탭 ➔ **`[About SMILES]`** ➔ **`[Check for Updates]`** 클릭
  - 또는 처음 설치할 때 썼던 PowerShell 명령어(`irm ... | iex`)를 다시 실행하면 최신 버전(`v1.7.0`)으로 자동 덮어쓰기 업그레이드됩니다.
* **깔끔하게 제거하기 (Uninstall):**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

---

## 📚 고급 정보 및 기업망 수동 설치 (클릭하여 펼치기)

::: details 🏢 사내 보안 정책으로 온라인 PowerShell 실행이 막혀 있는 경우 (오프라인 수동 설치)
1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases)에서 최신 `Smiles2Img-Office-x64.zip` (32비트 오피스는 `x86.zip`)과 `install.ps1`을 다운로드합니다.
2. 다운로드한 `.zip` 파일을 우클릭 ➔ **속성** ➔ 하단 **[차단 해제]**(Unblock)에 체크하고 **[확인]** 버튼을 누릅니다.
3. 압축을 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 폴더에 풉니다.
4. 해당 폴더에서 PowerShell을 열고 아래 명령어로 로컬 설치를 실행합니다:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```
:::

::: details 💻 지원 환경 및 시스템 요구사항
| 항목 | 지원 사양 | 비고 |
| :--- | :--- | :--- |
| **운영체제** | Windows 10 / Windows 11 | 32비트 및 64비트 모두 지원 |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | 데스크톱 설치 버전 |
| **필수 런타임** | .NET Framework 4.8 & Edge WebView2 | Windows 10/11에 기본 내장되어 별도 설치 불필요 |
:::

::: details ⚙️ 자동 설치 스크립트가 내부적으로 수행하는 작업
1. PC에 설치된 Office가 32비트(`x86`)인지 64비트(`x64`)인지 자동 판별합니다.
2. GitHub Releases에서 최신 배포 패키지를 받아 `%APPDATA%\Microsoft\AddIns\Smiles2Img` 경로에 배치합니다.
3. 다운로드된 파일의 Windows 보안 차단 표시(`Zone.Identifier`)를 자동 해제(`Unblock-File`)합니다.
4. 현재 사용자 계정(`HKCU`) 레지스트리에 Excel XLL 추가 기능과 Word/PowerPoint COM 추가 기능을 자동 등록합니다 (관리자 권한 불필요).
:::

::: details ✨ v1.7.0 주요 업데이트 요약
* **인셀 약물성(Lipinski) & 부분구조 스크리닝:** `=SMILES.MATCH`, `=SMILES.SIMILARITY`, `=SMILES2LIPINSKI`, `=SMILES2LOGP`, `=SMILES2TPSA` 추가
* **화합물 이름/CAS 번호 직접 변환:** `=NAME2IMG("Aspirin")`, `=NAME2SMILES("50-78-2")` 지원
* **문서 전체 일괄 스타일 변환 (`Batch Style`):** 클릭 한 번으로 컬러, 논문용 흑백(ACS), 흰색 배경, 다크 모드로 일괄 변경
* **무손실 벡터 복사 (`Copy SVG`) & 엑셀 `.sdf` 입출력:** Illustrator용 SVG 클립보드 복사 및 대용량 `.sdf` 표 가져오기/내보내기
* **개편된 분자 정보 창 (`Molecule Info`):** 한눈에 들어오는 물성 표와 엑셀 행(TSV)·마크다운·JSON 원클릭 내보내기
:::
