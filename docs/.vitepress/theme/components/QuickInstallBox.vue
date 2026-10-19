<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'

defineProps<{
  hideQuickstartBtn?: boolean
}>()

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const copied = ref(false)

const psCommand = 'irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex'

async function copyPsCommand() {
  try {
    await navigator.clipboard.writeText(psCommand)
    copied.value = true
    setTimeout(() => {
      copied.value = false
    }, 2000)
  } catch {
    // Fallback
  }
}
</script>

<template>
  <div class="s2i-widget qib-root">
    <div class="qib-main-card">
      <div class="qib-badge-row">
        <span class="qib-recommend-pill">
          {{ isKo ? '✨ 화학·바이오 연구원 권장 (초간단 설치)' : '✨ Recommended for Researchers (No Terminal Needed)' }}
        </span>
      </div>

      <h3 class="qib-title">
        {{ isKo ? '원클릭 자동 설치 파일로 10초 만에 시작하세요' : 'Start in 10 Seconds with the 1-Click Installer' }}
      </h3>
      <p class="qib-subtitle">
        {{
          isKo
            ? '복잡한 명령어 입력 없이, 아래 설치 파일을 내려받아 더블클릭하면 내 PC의 Excel · Word · PowerPoint(32/64비트)에 자동으로 등록됩니다.'
            : 'No command-line needed. Download the installer below and double-click it to automatically configure Excel, Word, and PowerPoint (32/64-bit).'
        }}
      </p>

      <div class="qib-cta-row">
        <a
          :href="withBase('/Install-func_SMILES2IMG.cmd')"
          download="Install-func_SMILES2IMG.cmd"
          class="qib-btn primary"
        >
          {{ isKo ? '📥 원클릭 자동 설치 파일 다운로드 (.cmd)' : '📥 Download 1-Click Installer (.cmd)' }}
        </a>
        <a
          v-if="!hideQuickstartBtn"
          :href="withBase(isKo ? '/ko/guide/getting-started' : '/guide/getting-started')"
          class="qib-btn secondary"
        >
          {{ isKo ? '🚀 설치 후 1분 첫걸음 가이드 ➔' : '🚀 Follow 1-Min Quickstart ➔' }}
        </a>
      </div>

      <div class="qib-trust-badges">
        <span class="qib-trust-item">
          ✅ {{ isKo ? '관리자 권한 불필요' : 'No Admin Rights Needed' }}
        </span>
        <span class="qib-trust-item">
          🛡️ {{ isKo ? '100% 오프라인 동작 (외부 서버 전송 없음)' : '100% Offline (Zero Cloud Calls)' }}
        </span>
        <span class="qib-trust-item">
          💻 {{ isKo ? 'Excel · Word · PPT 통합 설치' : 'Excel · Word · PPT All-in-One' }}
        </span>
        <span class="qib-trust-item">
          🎁 {{ isKo ? '무료 오픈소스 (MIT)' : '100% Free & Open Source' }}
        </span>
      </div>

      <div class="qib-steps">
        <div class="qib-step">
          <span class="qib-step-num">1</span>
          <span>
            {{
              isKo
                ? '실행 중인 Excel·Word·PPT를 닫고 내려받은 파일을 더블클릭하세요. (Windows 보호 안내가 뜨면 [추가 정보] → [실행] 클릭)'
                : 'Close open Office apps and double-click the downloaded file. (If Windows SmartScreen appears, click "More info" → "Run anyway".)'
            }}
          </span>
        </div>
        <div class="qib-step">
          <span class="qib-step-num">2</span>
          <span>
            {{
              isKo
                ? '엑셀을 열고 셀에 =NAME2IMG("Aspirin")을 입력하거나 Ctrl+Shift+K를 눌러 바로 구조식을 그려보세요!'
                : 'Open Excel and type =NAME2IMG("Aspirin") in any cell, or press Ctrl+Shift+K to draw!'
            }}
          </span>
        </div>
      </div>
    </div>

    <details class="qib-details">
      <summary>
        {{
          isKo
            ? '💻 개발자 또는 터미널 선호 유저용 (PowerShell 1줄 설치 / 오프라인 ZIP 패키지)'
            : '💻 For Developers & Terminal Users (1-Line PowerShell / Offline ZIP Package)'
        }}
      </summary>
      <div class="qib-details-body">
        <p class="qib-details-desc">
          {{
            isKo
              ? 'PowerShell 창에 아래 명령어를 붙여넣으면 최신 안정판(v1.7.1)이 즉시 설치 또는 업그레이드됩니다:'
              : 'Paste this one-liner into Windows PowerShell to install or upgrade to the latest release (v1.7.1):'
          }}
        </p>
        <div class="qib-code-row">
          <code>{{ psCommand }}</code>
          <button type="button" class="qib-copy-btn" @click="copyPsCommand">
            {{ copied ? (isKo ? '✓ 복사됨!' : '✓ Copied!') : (isKo ? '📋 명령어 복사' : '📋 Copy') }}
          </button>
        </div>
        <div class="qib-alt-links">
          <a href="https://github.com/naramdash/func_SMILES2IMG/releases/latest" target="_blank" rel="noopener">
            📦 {{ isKo ? 'GitHub Releases에서 오프라인 .zip 직접 다운로드 ➔' : 'Download Offline .zip from GitHub Releases ➔' }}
          </a>
          <a :href="withBase('/Uninstall-func_SMILES2IMG.cmd')" download="Uninstall-func_SMILES2IMG.cmd">
            🗑️ {{ isKo ? '원클릭 제거 파일 (.cmd)' : '1-Click Uninstaller (.cmd)' }}
          </a>
        </div>
      </div>
    </details>
  </div>
</template>

<style scoped>
.qib-root {
  padding: 24px;
  border-radius: 14px;
  background: var(--vp-c-bg-soft);
  border: 1px solid var(--vp-c-divider);
}
.qib-main-card {
  display: flex;
  flex-direction: column;
  gap: 14px;
}
.qib-recommend-pill {
  display: inline-block;
  font-size: 0.78rem;
  font-weight: 700;
  padding: 4px 12px;
  border-radius: 999px;
  background: var(--vp-c-brand-soft);
  color: var(--vp-c-brand-1);
}
.qib-title {
  margin: 0 !important;
  padding: 0 !important;
  border: none !important;
  font-size: 1.35rem;
  font-weight: 800;
  color: var(--vp-c-text-1);
  line-height: 1.3;
}
.qib-subtitle {
  margin: 0 !important;
  font-size: 0.95rem;
  color: var(--vp-c-text-2);
  line-height: 1.55;
}
.qib-cta-row {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  margin-top: 4px;
}
.qib-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 12px 22px;
  border-radius: 10px;
  font-size: 0.96rem;
  font-weight: 700;
  text-decoration: none !important;
  transition: all 0.15s ease;
  cursor: pointer;
}
.qib-btn.primary {
  background: var(--vp-c-brand-1);
  color: #ffffff !important;
  border: 1px solid var(--vp-c-brand-1);
  box-shadow: 0 4px 14px rgba(15, 118, 110, 0.25);
}
.dark .qib-btn.primary {
  color: #0f172a !important;
}
.qib-btn.primary:hover {
  background: var(--vp-c-brand-2);
  transform: translateY(-1px);
}
.qib-btn.secondary {
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1) !important;
  border: 1px solid var(--vp-c-divider);
}
.qib-btn.secondary:hover {
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1) !important;
}
.qib-trust-badges {
  display: flex;
  flex-wrap: wrap;
  gap: 8px 16px;
  padding-top: 4px;
}
.qib-trust-item {
  font-size: 0.82rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
}
.qib-steps {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 10px;
  margin-top: 4px;
}
.qib-step {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  background: var(--vp-c-bg);
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  padding: 10px 12px;
  font-size: 0.84rem;
  color: var(--vp-c-text-2);
  line-height: 1.45;
}
.qib-step-num {
  flex-shrink: 0;
  width: 22px;
  height: 22px;
  border-radius: 50%;
  background: var(--vp-c-brand-soft);
  color: var(--vp-c-brand-1);
  font-weight: 800;
  font-size: 0.78rem;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-top: 1px;
}
.qib-details {
  margin-top: 16px;
  padding-top: 14px;
  border-top: 1px solid var(--vp-c-divider);
}
.qib-details summary {
  cursor: pointer;
  font-size: 0.86rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
  user-select: none;
}
.qib-details summary:hover {
  color: var(--vp-c-brand-1);
}
.qib-details-body {
  margin-top: 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.qib-details-desc {
  margin: 0 !important;
  font-size: 0.84rem;
  color: var(--vp-c-text-2);
}
.qib-code-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  background: var(--vp-c-bg);
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  padding: 8px 12px;
  overflow-x: auto;
}
.qib-code-row code {
  font-size: 0.82rem;
  color: var(--vp-c-text-1);
  white-space: nowrap;
}
.qib-copy-btn {
  flex-shrink: 0;
  padding: 5px 12px;
  border-radius: 6px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
  color: var(--vp-c-text-1);
  font-size: 0.78rem;
  font-weight: 600;
  cursor: pointer;
}
.qib-copy-btn:hover {
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}
.qib-alt-links {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 0.82rem;
}
</style>
