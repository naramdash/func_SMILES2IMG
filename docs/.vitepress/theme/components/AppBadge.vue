<script setup lang="ts">
import { computed } from 'vue'
import { withBase } from 'vitepress'

const props = defineProps<{
  app: 'excel' | 'word' | 'powerpoint'
}>()

const meta = computed(() => {
  if (props.app === 'excel') {
    return { label: 'Excel', icon: '/icons/excel.svg', cls: 'is-excel' }
  }
  if (props.app === 'word') {
    return { label: 'Word', icon: '/icons/word.svg', cls: 'is-word' }
  }
  return { label: 'PowerPoint', icon: '/icons/powerpoint.svg', cls: 'is-ppt' }
})
</script>

<template>
  <span class="s2i-app-label vp-raw" :class="meta.cls">
    <img
      :src="withBase(meta.icon)"
      alt=""
      width="17"
      height="17"
      class="s2i-app-icon"
      aria-hidden="true"
    />
    <span>{{ meta.label }}</span>
  </span>
</template>

<style scoped>
.s2i-app-label {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  font-weight: 600;
  white-space: nowrap;
}

.s2i-app-icon {
  display: inline-block;
  width: 17px;
  height: 17px;
  flex-shrink: 0;
  margin: 0;
}

.is-excel {
  color: var(--s2i-excel);
}

.is-word {
  color: var(--s2i-word);
}

.is-ppt {
  color: var(--s2i-ppt);
}
</style>
