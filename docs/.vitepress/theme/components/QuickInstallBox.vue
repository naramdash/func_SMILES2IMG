<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData } from 'vitepress'

const props = withDefaults(
  defineProps<{
    variant?: 'hero' | 'page'
    hideQuickstartBtn?: boolean
  }>(),
  {
    variant: 'page',
    hideQuickstartBtn: false
  }
)

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const copied = ref(false)
const copiedUninstall = ref(false)

const psInstallCommand =
  'irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex'
const psUninstallCommand =
  'irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex'

async function copyText(text: string, isUninstall = false) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text)
    } else {
      const el = document.createElement('textarea')
      el.value = text
      el.style.position = 'fixed'
      el.style.left = '-9999px'
      el.style.top = '-9999px'
      document.body.appendChild(el)
      el.focus()
      el.select()
      document.execCommand('copy')
      document.body.removeChild(el)
    }

    if (isUninstall) {
      copiedUninstall.value = true
      setTimeout(() => {
        copiedUninstall.value = false
      }, 2000)
    } else {
      copied.value = true
      setTimeout(() => {
        copied.value = false
      }, 2000)
    }
  } catch (err) {
    console.error('Failed to copy', err)
  }
}
</script>

<template>
  <!-- Hero Variant: Compact terminal box directly under Hero action buttons -->
  <div v-if="variant === 'hero'" class="s2i-hero-qib">
    <div class="s2i-hero-qib-head">
      <span class="s2i-hero-qib-title">
        ⚡ {{ isKo ? '10초 PowerShell 자동 설치' : '10-Second Auto-Install (PowerShell)' }}
      </span>
      <span class="s2i-hero-qib-tag">
        {{ isKo ? '관리자 권한 불필요' : 'No Admin Rights' }}
      </span>
    </div>

    <div class="s2i-terminal">
      <div class="s2i-terminal-body">
        <span class="s2i-terminal-prompt">PS &gt;</span>
        <code class="s2i-terminal-code">{{ psInstallCommand }}</code>
      </div>
      <button
        type="button"
        class="s2i-terminal-copy-btn"
        :class="{ success: copied }"
        :aria-label="isKo ? '설치 명령어 복사' : 'Copy install command'"
        @click="copyText(psInstallCommand)"
      >
        <span class="s2i-copy-icon">{{ copied ? '✓' : '📋' }}</span>
        <span class="s2i-copy-label">{{ copied ? (isKo ? '복사됨!' : 'Copied!') : (isKo ? '복사' : 'Copy') }}</span>
      </button>
    </div>

    <div class="s2i-hero-qib-foot">
      <span>🛡️ {{ isKo ? '100% 로컬 오프라인' : '100% Local Offline' }}</span>
      <span class="s2i-dot">·</span>
      <span>💻 {{ isKo ? 'Excel · Word · PPT 자동 등록' : 'Auto-Registers Excel/Word/PPT' }}</span>
      <span class="s2i-dot">·</span>
      <span>🔄 {{ isKo ? '재실행 시 자동 업그레이드' : 'Re-run to Upgrade' }}</span>
    </div>
  </div>

  <!-- Page Variant: Used on Getting Started and Documentation pages -->
  <div v-else class="s2i-page-qib s2i-widget">
    <div class="s2i-page-qib-header">
      <div class="s2i-page-qib-badge">
        {{ isKo ? '⚡ 10초 PowerShell 자동 설치' : '⚡ 10-Second PowerShell Auto-Install' }}
      </div>
      <h3 class="s2i-page-qib-title">
        {{
          isKo
            ? 'Windows PowerShell에 명령어 한 줄만 붙여넣으면 끝납니다'
            : 'Paste One Line into Windows PowerShell — That’s It'
        }}
      </h3>
      <p class="s2i-page-qib-desc">
        {{
          isKo
            ? 'Excel, Word, PowerPoint(32비트/64비트)를 자동 감지하고 최신 안정판 바이너리 다운로드, Windows 보안 해제(Unblock), 추가 기능 등록을 10초 만에 완료합니다.'
            : 'Auto-detects Excel, Word, and PowerPoint (32/64-bit), downloads the latest release, unblocks security flags, and registers add-ins in 10 seconds.'
        }}
      </p>
    </div>

    <!-- Terminal Command Box -->
    <div class="s2i-page-terminal">
      <div class="s2i-page-terminal-bar">
        <div class="s2i-page-terminal-dots">
          <span class="dot red"></span>
          <span class="dot yellow"></span>
          <span class="dot green"></span>
        </div>
        <span class="s2i-page-terminal-title">Windows PowerShell</span>
      </div>

      <div class="s2i-page-terminal-content">
        <div class="s2i-page-terminal-cmd">
          <span class="s2i-terminal-prompt">PS &gt;</span>
          <code class="s2i-terminal-code">{{ psInstallCommand }}</code>
        </div>
        <button
          type="button"
          class="s2i-page-copy-btn"
          :class="{ success: copied }"
          @click="copyText(psInstallCommand)"
        >
          <span class="s2i-copy-icon">{{ copied ? '✓' : '📋' }}</span>
          <span class="s2i-copy-label">
            {{
              copied
                ? isKo
                  ? '클립보드에 복사 완료!'
                  : 'Copied to Clipboard!'
                : isKo
                  ? '명령어 복사'
                  : 'Copy Command'
            }}
          </span>
        </button>
      </div>
    </div>

    <!-- 3 Simple Steps -->
    <div class="s2i-page-steps">
      <div class="s2i-step-card">
        <div class="s2i-step-num">1</div>
        <div class="s2i-step-body">
          <strong>{{ isKo ? '오피스 닫기' : 'Close Office' }}</strong>
          <p>
            {{
              isKo
                ? '실행 중인 Excel, Word, PowerPoint 창을 완전히 닫습니다.'
                : 'Close any running Excel, Word, or PowerPoint windows.'
            }}
          </p>
        </div>
      </div>

      <div class="s2i-step-card">
        <div class="s2i-step-num">2</div>
        <div class="s2i-step-body">
          <strong>{{ isKo ? '명령어 실행' : 'Paste & Run' }}</strong>
          <p>
            {{
              isKo
                ? 'Windows 시작 메뉴에서 PowerShell을 열고 위 명령어를 붙여넣어 실행합니다.'
                : 'Open PowerShell from the Start menu, paste the command, and press Enter.'
            }}
          </p>
        </div>
      </div>

      <div class="s2i-step-card">
        <div class="s2i-step-num">3</div>
        <div class="s2i-step-body">
          <strong>{{ isKo ? '오피스 재시작' : 'Restart Office' }}</strong>
          <p>
            {{
              isKo
                ? '오피스를 다시 열면 상단 리본에 [SMILES] 탭이 바로 나타납니다!'
                : 'Reopen Office to see the [SMILES] tab on your top ribbon menu!'
            }}
          </p>
        </div>
      </div>
    </div>

    <!-- Trust Badges -->
    <div class="s2i-trust-pills">
      <span class="s2i-trust-pill">
        ✅ {{ isKo ? '관리자 권한 불필요 (일반 계정 실행)' : 'No Admin Rights Required' }}
      </span>
      <span class="s2i-trust-pill">
        🛡️ {{ isKo ? '100% 로컬 오프라인 (연구 기밀 보호)' : '100% Local & Offline' }}
      </span>
      <span class="s2i-trust-pill">
        💻 {{ isKo ? 'Excel · Word · PowerPoint 통합 설치' : 'Excel · Word · PPT Suite' }}
      </span>
      <span class="s2i-trust-pill">
        🎁 {{ isKo ? '무료 오픈소스 (MIT 라이선스)' : 'Free & Open Source (MIT)' }}
      </span>
    </div>

    <!-- Expandable Advanced Info: Offline ZIP and Clean Uninstall -->
    <details class="s2i-page-details">
      <summary class="s2i-details-summary">
        {{
          isKo
            ? '📦 오프라인 사내 폐쇄망 설치 또는 삭제(제거) 방법'
            : '📦 Offline / Air-Gapped Network Install & Clean Uninstall'
        }}
      </summary>
      <div class="s2i-details-content">
        <div class="s2i-details-sub">
          <strong>📦 {{ isKo ? '인터넷 차단망 수동 설치:' : 'Offline Manual ZIP Install:' }}</strong>
          <p>
            {{
              isKo
                ? '외부 인터넷 접속이 차단된 사내 환경에서는 GitHub Releases에서 ZIP 패키지를 내려받아 압축 해제 후 install.ps1을 직접 실행하시면 됩니다.'
                : 'For air-gapped corporate PCs, download the ZIP package from GitHub Releases, extract it, and execute install.ps1 locally.'
            }}
          </p>
          <a
            href="https://github.com/naramdash/func_SMILES2IMG/releases/latest"
            target="_blank"
            rel="noopener"
            class="s2i-details-link"
          >
            📦 {{ isKo ? 'GitHub Releases 최신 버전 패키지 다운로드 ➔' : 'Download Latest Packages from GitHub Releases ➔' }}
          </a>
        </div>

        <div class="s2i-details-sub">
          <strong>🗑️ {{ isKo ? '깔끔한 원클릭 제거 (Uninstall):' : 'Clean 1-Line Uninstall:' }}</strong>
          <p>
            {{
              isKo
                ? 'PowerShell에 아래 한 줄을 실행하면 추가 기능 등록 및 관련 파일이 완전히 삭제됩니다:'
                : 'Run this single command in PowerShell to cleanly remove all add-in files and registry entries:'
            }}
          </p>
          <div class="s2i-uninstall-row">
            <code class="s2i-uninstall-code">{{ psUninstallCommand }}</code>
            <button
              type="button"
              class="s2i-uninstall-copy-btn"
              :class="{ success: copiedUninstall }"
              @click="copyText(psUninstallCommand, true)"
            >
              {{ copiedUninstall ? (isKo ? '✓ 복사됨!' : '✓ Copied!') : (isKo ? '제거 명령어 복사' : 'Copy Uninstall') }}
            </button>
          </div>
        </div>
      </div>
    </details>
  </div>
</template>

<style scoped>
/* ============================================================
   HERO VARIANT (Ultra-compact, placed below Hero action buttons)
   ============================================================ */
.s2i-hero-qib {
  margin-top: 18px;
  width: 100%;
  max-width: 540px;
  background: var(--vp-c-bg-soft);
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  padding: 12px 14px;
  box-shadow: 0 4px 16px rgba(15, 23, 42, 0.06);
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.s2i-hero-qib-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}

.s2i-hero-qib-title {
  font-size: 0.84rem;
  font-weight: 700;
  color: var(--vp-c-text-1);
  display: inline-flex;
  align-items: center;
  gap: 4px;
}

.s2i-hero-qib-tag {
  font-size: 0.72rem;
  font-weight: 600;
  color: var(--vp-c-brand-1);
  background: var(--vp-c-brand-soft);
  padding: 2px 8px;
  border-radius: 999px;
  white-space: nowrap;
}

.s2i-terminal {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  background: #0f172a;
  border: 1px solid #334155;
  border-radius: 8px;
  padding: 6px 10px;
}

.s2i-terminal-body {
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
  flex: 1;
  overflow-x: auto;
  scrollbar-width: none;
}
.s2i-terminal-body::-webkit-scrollbar {
  display: none;
}

.s2i-terminal-prompt {
  font-family: var(--vp-font-family-mono);
  font-size: 0.78rem;
  font-weight: 800;
  color: #38bdf8;
  user-select: none;
  flex-shrink: 0;
}

.s2i-terminal-code {
  font-family: var(--vp-font-family-mono);
  font-size: 0.78rem;
  color: #f1f5f9;
  white-space: nowrap;
  background: transparent !important;
  padding: 0 !important;
  margin: 0 !important;
  user-select: all;
}

.s2i-terminal-copy-btn {
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  background: var(--vp-c-brand-1);
  color: #ffffff;
  border: none;
  border-radius: 6px;
  padding: 5px 12px;
  font-size: 0.78rem;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.15s ease;
  user-select: none;
}
.s2i-terminal-copy-btn:hover {
  background: var(--vp-c-brand-2);
  transform: translateY(-1px);
}
.s2i-terminal-copy-btn.success {
  background: #10b981;
}

.s2i-hero-qib-foot {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px 8px;
  font-size: 0.74rem;
  font-weight: 500;
  color: var(--vp-c-text-2);
  line-height: 1.4;
}

.s2i-dot {
  opacity: 0.4;
}

/* ============================================================
   PAGE VARIANT (Full width, rich step cards for Getting Started)
   ============================================================ */
.s2i-page-qib {
  padding: 24px;
  border-radius: 14px;
  background: var(--vp-c-bg-soft);
  border: 1px solid var(--vp-c-divider);
  display: flex;
  flex-direction: column;
  gap: 18px;
  margin: 18px 0 28px 0;
}

.s2i-page-qib-badge {
  display: inline-block;
  font-size: 0.78rem;
  font-weight: 700;
  padding: 4px 12px;
  border-radius: 999px;
  background: var(--vp-c-brand-soft);
  color: var(--vp-c-brand-1);
  margin-bottom: 8px;
}

.s2i-page-qib-title {
  margin: 0 0 6px 0 !important;
  padding: 0 !important;
  border: none !important;
  font-size: 1.35rem;
  font-weight: 800;
  color: var(--vp-c-text-1);
  line-height: 1.3;
}

.s2i-page-qib-desc {
  margin: 0 !important;
  font-size: 0.94rem;
  color: var(--vp-c-text-2);
  line-height: 1.55;
}

/* Terminal box inside page */
.s2i-page-terminal {
  background: #0f172a;
  border: 1px solid #334155;
  border-radius: 10px;
  overflow: hidden;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
}

.s2i-page-terminal-bar {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 14px;
  background: #1e293b;
  border-bottom: 1px solid #334155;
}

.s2i-page-terminal-dots {
  display: flex;
  gap: 6px;
}

.s2i-page-terminal-dots .dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
}
.s2i-page-terminal-dots .dot.red {
  background: #ef4444;
}
.s2i-page-terminal-dots .dot.yellow {
  background: #f59e0b;
}
.s2i-page-terminal-dots .dot.green {
  background: #10b981;
}

.s2i-page-terminal-title {
  font-size: 0.74rem;
  font-weight: 600;
  color: #94a3b8;
  font-family: var(--vp-font-family-mono);
}

.s2i-page-terminal-content {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  padding: 12px 14px;
}

.s2i-page-terminal-cmd {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
  flex: 1;
  overflow-x: auto;
}

.s2i-page-copy-btn {
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: var(--vp-c-brand-1);
  color: #ffffff !important;
  border: none;
  border-radius: 8px;
  padding: 8px 16px;
  font-size: 0.88rem;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.15s ease;
  user-select: none;
  box-shadow: 0 2px 8px rgba(15, 118, 110, 0.3);
}
.s2i-page-copy-btn:hover {
  background: var(--vp-c-brand-2);
  transform: translateY(-1px);
}
.s2i-page-copy-btn.success {
  background: #10b981;
}

/* 3 Step Cards */
.s2i-page-steps {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 12px;
}

.s2i-step-card {
  display: flex;
  align-items: flex-start;
  gap: 12px;
  background: var(--vp-c-bg);
  border: 1px solid var(--vp-c-divider);
  border-radius: 10px;
  padding: 12px 14px;
}

.s2i-step-num {
  flex-shrink: 0;
  width: 26px;
  height: 26px;
  border-radius: 50%;
  background: var(--vp-c-brand-soft);
  color: var(--vp-c-brand-1);
  font-weight: 800;
  font-size: 0.84rem;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-top: 2px;
}

.s2i-step-body {
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.s2i-step-body strong {
  font-size: 0.88rem;
  color: var(--vp-c-text-1);
}
.s2i-step-body p {
  margin: 0 !important;
  font-size: 0.82rem;
  color: var(--vp-c-text-2);
  line-height: 1.45;
}

/* Trust Badges */
.s2i-trust-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 8px 14px;
}

.s2i-trust-pill {
  font-size: 0.82rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
}

/* Details Section */
.s2i-page-details {
  border-top: 1px solid var(--vp-c-divider);
  padding-top: 12px;
  margin-top: 4px;
}

.s2i-details-summary {
  cursor: pointer;
  font-size: 0.86rem;
  font-weight: 700;
  color: var(--vp-c-text-2);
  user-select: none;
}
.s2i-details-summary:hover {
  color: var(--vp-c-brand-1);
}

.s2i-details-content {
  margin-top: 12px;
  display: flex;
  flex-direction: column;
  gap: 14px;
  padding: 10px 14px;
  background: var(--vp-c-bg);
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
}

.s2i-details-sub {
  display: flex;
  flex-direction: column;
  gap: 4px;
  font-size: 0.84rem;
  color: var(--vp-c-text-2);
}
.s2i-details-sub p {
  margin: 0 !important;
  line-height: 1.45;
}
.s2i-details-link {
  font-weight: 600;
  color: var(--vp-c-brand-1) !important;
  text-decoration: underline;
  margin-top: 2px;
}

.s2i-uninstall-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  background: var(--vp-c-bg-soft);
  border: 1px solid var(--vp-c-divider);
  border-radius: 6px;
  padding: 6px 10px;
  margin-top: 4px;
  overflow-x: auto;
}

.s2i-uninstall-code {
  font-size: 0.78rem;
  font-family: var(--vp-font-family-mono);
  color: var(--vp-c-text-1);
  white-space: nowrap;
}

.s2i-uninstall-copy-btn {
  flex-shrink: 0;
  padding: 4px 10px;
  border-radius: 5px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font-size: 0.74rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.15s ease;
}
.s2i-uninstall-copy-btn:hover {
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}
.s2i-uninstall-copy-btn.success {
  background: #10b981;
  color: #ffffff;
  border-color: #10b981;
}

/* Mobile responsive adjustments */
@media (max-width: 640px) {
  .s2i-terminal {
    flex-direction: column;
    align-items: stretch;
  }
  .s2i-terminal-copy-btn {
    justify-content: center;
  }
  .s2i-page-terminal-content {
    flex-direction: column;
    align-items: stretch;
  }
  .s2i-page-copy-btn {
    justify-content: center;
  }
  .s2i-uninstall-row {
    flex-direction: column;
    align-items: stretch;
  }
}
</style>
