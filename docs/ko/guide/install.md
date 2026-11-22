# 설치 (Installation)

**관리자 권한 없이 10초 만에** Microsoft Excel, Word, PowerPoint 전체에 `func_SMILES2IMG`를 설치합니다.

## 1. 10초 PowerShell 자동 설치

1. 실행 중인 Excel, Word, PowerPoint 창을 모두 닫습니다.
2. Windows 시작 메뉴에서 **PowerShell**을 열고 아래 한 줄을 붙여넣어 실행합니다:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip `v1.8.1` 정식 포함 기능
위 명령어를 실행하면 Office 비트(32비트/64비트)를 자동 감지해 `func_SMILES2IMG v1.8.1` 패키지를 내려받고 파일 차단 해제(`Unblock-File`) 및 HKCU 레지스트리 등록까지 자동 완료합니다. Excel·Word·PowerPoint 통합 `Chemical Structure >` 우클릭 메뉴, 40개 엑셀 함수, Ketcher 2D, Mol\* 3D, 12열 `.sdf` 가져오기가 모두 포함됩니다.
:::

## 2. 오피스 재시작 & `[SMILES]` 리본 탭 확인

1. 열려 있던 Excel, Word, PowerPoint를 **모두 닫고 다시 실행**합니다(추가 기능은 오피스 시작 시 로드됩니다).
2. 상단 리본에 **`[SMILES]`** 탭이 나타났는지 확인합니다.
3. 언제든 **`[About func_SMILES2IMG]`** 버튼을 클릭해 설치된 버전, 서드파티 라이선스 고지, 업데이트 여부를 확인할 수 있습니다.

<FeatureMedia id="common-about-dialog" />

## 3. 사내 보안망(오프라인) 수동 ZIP 설치 {#offline-zip}

사내 방화벽 등으로 PowerShell 온라인 다운로드가 제한된 경우 수동으로 설치할 수 있습니다:

1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest)에서 `func_SMILES2IMG-x64.zip` (32비트 오피스는 `x86`)을 내려받습니다.
2. `.zip` 파일 우클릭 ➔ **속성** ➔ 하단 **[차단 해제]** 체크 ➔ **[확인]** 클릭.
3. 압축을 푼 폴더에서 PowerShell을 열고 아래 명령어를 실행합니다:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

## 4. 설치 시 보안 경고 해결

::: details 🛡️ Windows 11 스마트 앱 컨트롤(SAC) 차단 또는 설치 후 `#NAME?` 오류
Windows 11의 **스마트 앱 컨트롤**이 강제 켜짐 상태면 서명되지 않은 오픈소스 `.xll` 추가 기능 실행을 차단할 수 있습니다. 이를 꺼도 **Microsoft Defender 백신(실시간 바이러스 감시)은 100% 그대로 유지**됩니다.
1. Windows **설정** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **개인 정보 및 보안** ➔ **Windows 보안**
2. **앱 및 브라우저 컨트롤** ➔ **스마트 앱 컨트롤 설정**
3. **꺼짐(Off)** 상태로 변경한 뒤 오피스를 다시 시작합니다.
:::

::: details 🔓 수동 다운로드한 `.xll` 파일이 엑셀에서 차단될 때
`.xll` 파일 우클릭 ➔ **속성** ➔ 하단 **[차단 해제]**(Unblock) 체크 후 **[확인]** 버튼을 누르거나, 차단 해제를 자동으로 처리하는 상단 PowerShell 1줄 설치 명령어를 실행하세요.
:::

## 5. 업데이트 및 깔끔한 제거

* **최신 버전으로 업데이트:** 리본 **`[SMILES]` ➔ `[About func_SMILES2IMG]` ➔ `[Check for Updates]`** 버튼을 누르거나, 상단 설치 명령어를 다시 실행하세요.
* **완전 제거 (파일 및 레지스트리 항목 모두 정리):**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

## 다음 단계

- [**시작하기: 1분 만에 첫 분자 그리기**](/ko/guide/getting-started) — Excel 첫 수식 입력, Word 인라인 삽입, PowerPoint 3D 슬라이드를 바로 따라 해보세요.
- [**자주 묻는 질문 (FAQ)**](/ko/guide/faq) — #NAME? · #VALUE! 해결, 구형 Excel(.FLOAT), 문서 공유, SVG 내보내기 안내
