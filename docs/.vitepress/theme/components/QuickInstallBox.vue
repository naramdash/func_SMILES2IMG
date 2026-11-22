<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData } from 'vitepress'
import { copyToClipboard } from '../clipboard'

defineProps<{
  inline?: boolean
}>()

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const copied = ref(false)
let copiedTimer: ReturnType<typeof setTimeout> | undefined

const psInstallCommand =
  'irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex'

const copyLabel = computed(() =>
  copied.value
    ? (isKo.value ? '복사됨' : 'Copied')
    : (isKo.value ? '설치 명령어 복사' : 'Copy install command')
)

async function copyInstall() {
  if (await copyToClipboard(psInstallCommand)) {
    copied.value = true
    if (copiedTimer) clearTimeout(copiedTimer)
    copiedTimer = setTimeout(() => {
      copied.value = false
    }, 2000)
  }
}
</script>

<template>
  <div class="s2i-hero-qib-wrap vp-raw" :class="{ 'is-inline': inline }">
    <div class="s2i-hero-qib">
      <div class="s2i-terminal-body">
        <span class="s2i-ps-brand">
          <Icon icon="vscode-icons:file-type-powershell" width="20" height="20" class="s2i-ps-icon" />
          <span class="s2i-terminal-label">PowerShell</span>
        </span>
        <span class="s2i-terminal-prompt">&gt;</span>
        <code class="s2i-terminal-code">{{ psInstallCommand }}</code>
      </div>
      <button
        type="button"
        class="s2i-terminal-copy-btn"
        :class="{ success: copied }"
        :aria-label="copyLabel"
        :title="copyLabel"
        @click="copyInstall"
      >
        <Icon :icon="copied ? 'lucide:check' : 'lucide:copy'" width="16" height="16" aria-hidden="true" />
      </button>
    </div>
  </div>
</template>

<style scoped>
.s2i-hero-qib-wrap {
  padding: 0 24px;
  margin-bottom: 16px;
}

.s2i-hero-qib-wrap.is-inline {
  padding: 0;
  margin: 12px 0 16px;
}

.s2i-hero-qib-wrap.is-inline .s2i-hero-qib {
  max-width: none;
}

@media (min-width: 640px) {
  .s2i-hero-qib-wrap {
    padding: 0 48px;
  }
}

@media (min-width: 960px) {
  .s2i-hero-qib-wrap {
    padding: 0 64px;
  }
}

.s2i-hero-qib {
  max-width: 1152px;
  margin: 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  background: var(--vp-c-bg-soft);
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  padding: 8px 8px 8px 16px;
}

.s2i-terminal-body {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
  flex: 1;
  overflow-x: auto;
  scrollbar-width: none;
}
.s2i-terminal-body::-webkit-scrollbar {
  display: none;
}

.s2i-ps-brand {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  flex-shrink: 0;
  user-select: none;
}

.s2i-ps-icon {
  display: block;
  flex-shrink: 0;
}

.s2i-terminal-label {
  font-family: var(--vp-font-family-mono);
  font-size: 0.8rem;
  font-weight: 700;
  color: var(--vp-c-text-2);
  user-select: none;
  flex-shrink: 0;
}

.s2i-terminal-prompt {
  font-family: var(--vp-font-family-mono);
  font-size: 0.86rem;
  font-weight: 700;
  color: var(--vp-c-brand-1);
  user-select: none;
  flex-shrink: 0;
}

.s2i-terminal-code {
  font-family: var(--vp-font-family-mono);
  font-size: 0.85rem;
  color: var(--vp-c-text-1);
  white-space: nowrap;
  background: transparent;
  padding: 0;
  margin: 0;
  user-select: all;
}

.s2i-terminal-copy-btn {
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 34px;
  height: 34px;
  padding: 0;
  background: var(--vp-c-brand-1);
  color: #ffffff;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  transition: background-color 0.15s ease;
  user-select: none;
}
.s2i-terminal-copy-btn:hover {
  background: var(--vp-c-brand-2);
}
.dark .s2i-terminal-copy-btn {
  color: #0f172a;
}
.s2i-terminal-copy-btn.success {
  background: #10b981;
}
</style>
