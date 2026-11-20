<script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import { useData } from 'vitepress'
import { useSidebar } from 'vitepress/theme'
import { Icon } from '@iconify/vue'

const STORAGE_KEY = 's2i-sidebar-collapsed'
const CLASS_NAME = 's2i-sidebar-collapsed'

const { lang } = useData()
const { hasSidebar } = useSidebar()
const isKo = computed(() => lang.value.startsWith('ko'))
const collapsed = ref(false)

const label = computed(() => {
  if (isKo.value) {
    return collapsed.value ? '사이드바 열기' : '사이드바 접기'
  }
  return collapsed.value ? 'Show sidebar' : 'Hide sidebar'
})

function applyState(next: boolean) {
  collapsed.value = next
  if (typeof document !== 'undefined') {
    document.documentElement.classList.toggle(CLASS_NAME, next)
  }
}

function toggle() {
  const next = !collapsed.value
  applyState(next)
  try {
    localStorage.setItem(STORAGE_KEY, next ? '1' : '0')
  } catch {
    // ignore storage access errors
  }
}

onMounted(() => {
  try {
    if (localStorage.getItem(STORAGE_KEY) === '1') {
      applyState(true)
    }
  } catch {
    // ignore storage access errors
  }
})
</script>

<template>
  <button
    v-if="hasSidebar"
    type="button"
    class="s2i-sidebar-toggle"
    :aria-label="label"
    :aria-expanded="!collapsed"
    :title="label"
    @click="toggle"
  >
    <Icon
      :icon="collapsed ? 'lucide:panel-left-open' : 'lucide:panel-left-close'"
      width="18"
      height="18"
      aria-hidden="true"
    />
  </button>
</template>

<style scoped>
.s2i-sidebar-toggle {
  display: none;
  align-items: center;
  justify-content: center;
  width: 36px;
  height: 36px;
  margin-left: 24px;
  border-radius: 8px;
  border: 1px solid transparent;
  background: transparent;
  color: var(--vp-c-text-2);
  cursor: pointer;
  flex-shrink: 0;
  transition: color 0.2s, background-color 0.2s, border-color 0.2s;
}

.s2i-sidebar-toggle:hover {
  color: var(--vp-c-text-1);
  background: var(--vp-c-bg-alt);
  border-color: var(--vp-c-divider);
}

@media (min-width: 960px) {
  .s2i-sidebar-toggle {
    display: inline-flex;
  }
}
</style>
