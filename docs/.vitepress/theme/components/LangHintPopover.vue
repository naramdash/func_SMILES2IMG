<script setup lang="ts">
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import { useData, withBase } from 'vitepress'

const STORAGE_KEY = 's2i-lang-hint-dismissed'
const ACTIVE_CLASS = 's2i-lang-hint-active'

const { lang, page } = useData()

const visible = ref(false)
const inView = ref(true)
const targetLocale = ref<'ko' | 'en'>('ko')
const popoverRef = ref<HTMLElement | null>(null)
const pos = ref({ top: 56, left: 0, arrowLeft: 120 })

let showTimer: ReturnType<typeof setTimeout> | null = null
let resizeObserver: ResizeObserver | null = null
let navEl: HTMLElement | null = null
let translationsEl: HTMLElement | null = null

const isPageKo = computed(() => lang.value.startsWith('ko'))

const copy = computed(() => {
  if (targetLocale.value === 'ko') {
    return {
      title: '한국어로 보시겠습니까?',
      switchLabel: '한국어로 보기',
      stayLabel: 'English 유지',
      closeLabel: '언어 안내 닫기',
      dialogLabel: '한국어 문서 전환 안내'
    }
  }
  return {
    title: 'View this page in English?',
    switchLabel: 'Switch to English',
    stayLabel: 'Stay in 한국어',
    closeLabel: 'Close language prompt',
    dialogLabel: 'Switch to English documentation'
  }
})

const targetHref = computed(() => {
  const rel = (page.value.relativePath || '')
    .replace(/(^|\/)index\.md$/, '$1')
    .replace(/\.md$/, '')
  const hash = typeof window !== 'undefined' ? window.location.hash : ''

  if (targetLocale.value === 'ko') {
    const clean = rel.replace(/^ko\//, '')
    return withBase('/ko/' + clean) + hash
  }
  const clean = rel.replace(/^ko\//, '')
  return withBase('/' + clean) + hash
})

function setActiveClass(active: boolean) {
  if (typeof document !== 'undefined') {
    document.documentElement.classList.toggle(ACTIVE_CLASS, active)
  }
}

function isDismissed(): boolean {
  try {
    return sessionStorage.getItem(STORAGE_KEY) === '1'
  } catch {
    return false
  }
}

function dismiss() {
  visible.value = false
  setActiveClass(false)
  if (showTimer) {
    clearTimeout(showTimer)
    showTimer = null
  }
  try {
    sessionStorage.setItem(STORAGE_KEY, '1')
  } catch {
    // ignore storage access errors
  }
}

function updatePosition() {
  if (!visible.value || typeof window === 'undefined') return
  const btn = document.querySelector<HTMLElement>('.VPNavBarTranslations .button')
  if (!btn) {
    inView.value = false
    return
  }

  const rect = btn.getBoundingClientRect()
  if (rect.width === 0 || rect.bottom <= 0 || rect.top > window.innerHeight) {
    inView.value = false
    return
  }

  inView.value = true
  const anchorX = rect.left + rect.width / 2
  const top = Math.round(rect.top + rect.height / 2 + 24)
  const popW = popoverRef.value?.offsetWidth || 232
  const margin = 12
  const left = Math.round(
    Math.max(margin, Math.min(rect.right - popW, window.innerWidth - popW - margin))
  )
  const arrowLeft = Math.round(Math.max(18, Math.min(anchorX - left, popW - 18)))

  pos.value = { top, left, arrowLeft }
}

function onKeydown(e: KeyboardEvent) {
  if (e.key === 'Escape' && visible.value) {
    dismiss()
  }
}

function onTranslationsInteract() {
  if (visible.value) {
    dismiss()
  }
}

function detectMismatch(): 'ko' | 'en' | null {
  if (typeof navigator === 'undefined') return null
  const primary = (navigator.languages?.[0] || navigator.language || '').toLowerCase()
  if (!primary) return null

  const browserIsKo = primary.startsWith('ko')
  if (browserIsKo && !isPageKo.value) return 'ko'
  if (!browserIsKo && isPageKo.value) return 'en'
  return null
}

watch(
  () => page.value.relativePath,
  () => {
    if (!visible.value) return
    if (!detectMismatch()) {
      dismiss()
    } else {
      nextTick(() => requestAnimationFrame(updatePosition))
    }
  }
)

onMounted(() => {
  if (isDismissed()) return
  const suggested = detectMismatch()
  if (!suggested) return

  targetLocale.value = suggested

  showTimer = setTimeout(() => {
    if (isDismissed() || !detectMismatch()) return
    visible.value = true
    setActiveClass(true)

    nextTick(() => {
      updatePosition()
      requestAnimationFrame(updatePosition)
    })

    window.addEventListener('resize', updatePosition, { passive: true })
    window.addEventListener('scroll', updatePosition, { passive: true })
    window.addEventListener('keydown', onKeydown)

    translationsEl = document.querySelector<HTMLElement>('.VPNavBarTranslations')
    if (translationsEl) {
      translationsEl.addEventListener('mouseenter', onTranslationsInteract)
      translationsEl.addEventListener('click', onTranslationsInteract)
    }

    navEl = document.querySelector<HTMLElement>('.VPNavBar')
    if (navEl && typeof ResizeObserver !== 'undefined') {
      resizeObserver = new ResizeObserver(updatePosition)
      resizeObserver.observe(navEl)
    }
  }, 450)
})

onUnmounted(() => {
  setActiveClass(false)
  if (showTimer) clearTimeout(showTimer)
  if (typeof window !== 'undefined') {
    window.removeEventListener('resize', updatePosition)
    window.removeEventListener('scroll', updatePosition)
    window.removeEventListener('keydown', onKeydown)
  }
  if (translationsEl) {
    translationsEl.removeEventListener('mouseenter', onTranslationsInteract)
    translationsEl.removeEventListener('click', onTranslationsInteract)
  }
  if (resizeObserver) {
    resizeObserver.disconnect()
  }
})
</script>

<template>
  <Transition name="s2i-lang-hint">
    <div
      v-if="visible && inView"
      ref="popoverRef"
      class="s2i-lang-hint"
      role="dialog"
      :aria-label="copy.dialogLabel"
      :style="{
        top: `${pos.top}px`,
        left: `${pos.left}px`,
        '--s2i-arrow-left': `${pos.arrowLeft}px`
      }"
    >
      <span class="s2i-lang-hint-arrow" aria-hidden="true" />
      <div class="s2i-lang-hint-head">
        <span class="s2i-lang-hint-dot" aria-hidden="true" />
        <span class="s2i-lang-hint-title">{{ copy.title }}</span>
        <button
          type="button"
          class="s2i-lang-hint-close"
          :aria-label="copy.closeLabel"
          @click="dismiss"
        >
          <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M18 6 6 18M6 6l12 12" />
          </svg>
        </button>
      </div>
      <div class="s2i-lang-hint-actions">
        <a
          class="s2i-lang-hint-primary"
          :href="targetHref"
          @click="dismiss"
        >
          {{ copy.switchLabel }}
        </a>
        <button
          type="button"
          class="s2i-lang-hint-secondary"
          @click="dismiss"
        >
          {{ copy.stayLabel }}
        </button>
      </div>
    </div>
  </Transition>
</template>

<style scoped>
.s2i-lang-hint {
  position: fixed;
  z-index: 60;
  width: max-content;
  max-width: calc(100vw - 24px);
  padding: 11px 13px 12px;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  background-color: var(--vp-c-bg-elv);
  box-shadow: var(--vp-shadow-3);
  color: var(--vp-c-text-1);
}

.s2i-lang-hint-arrow {
  position: absolute;
  top: -6px;
  left: var(--s2i-arrow-left, 50%);
  width: 10px;
  height: 10px;
  transform: translateX(-50%) rotate(45deg);
  background-color: var(--vp-c-bg-elv);
  border-top: 1px solid var(--vp-c-divider);
  border-left: 1px solid var(--vp-c-divider);
}

.s2i-lang-hint-head {
  display: flex;
  align-items: center;
  gap: 7px;
  margin-bottom: 10px;
}

.s2i-lang-hint-dot {
  width: 7px;
  height: 7px;
  border-radius: 9999px;
  background-color: #ef4444;
  flex-shrink: 0;
}

.s2i-lang-hint-title {
  font-size: 13px;
  font-weight: 600;
  line-height: 1.35;
  color: var(--vp-c-text-1);
  white-space: nowrap;
}

.s2i-lang-hint-close {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-left: auto;
  padding: 3px;
  border-radius: 6px;
  color: var(--vp-c-text-3);
  cursor: pointer;
  transition: color 0.15s ease, background-color 0.15s ease;
}

.s2i-lang-hint-close:hover {
  color: var(--vp-c-text-1);
  background-color: var(--vp-c-default-soft);
}

.s2i-lang-hint-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.s2i-lang-hint-primary {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 5px 12px;
  border-radius: 8px;
  font-size: 12.5px;
  font-weight: 600;
  line-height: 1.3;
  text-decoration: none;
  white-space: nowrap;
  color: var(--vp-button-brand-text);
  background-color: var(--vp-button-brand-bg);
  transition: background-color 0.2s ease;
}

.s2i-lang-hint-primary:hover {
  background-color: var(--vp-button-brand-hover-bg);
}

.s2i-lang-hint-secondary {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 5px 10px;
  border-radius: 8px;
  font-size: 12.5px;
  font-weight: 500;
  line-height: 1.3;
  white-space: nowrap;
  color: var(--vp-c-text-2);
  background-color: var(--vp-c-default-soft);
  cursor: pointer;
  transition: color 0.2s ease, background-color 0.2s ease;
}

.s2i-lang-hint-secondary:hover {
  color: var(--vp-c-text-1);
  background-color: var(--vp-button-alt-hover-bg);
}

.s2i-lang-hint-enter-active,
.s2i-lang-hint-leave-active {
  transition: opacity 0.2s ease, transform 0.2s ease;
}

.s2i-lang-hint-enter-from,
.s2i-lang-hint-leave-to {
  opacity: 0;
  transform: translateY(-4px);
}

@media (prefers-reduced-motion: reduce) {
  .s2i-lang-hint-enter-active,
  .s2i-lang-hint-leave-active {
    transition: opacity 0.1s ease;
  }
  .s2i-lang-hint-enter-from,
  .s2i-lang-hint-leave-to {
    transform: none;
  }
}
</style>
