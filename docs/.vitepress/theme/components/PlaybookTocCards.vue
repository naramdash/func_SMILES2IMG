<script setup lang="ts">
import { computed } from 'vue'
import { useData, withBase } from 'vitepress'

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const cards = computed(() => [
  {
    id: 'excel',
    cls: 's2i-pb-toc-excel',
    name: 'Excel',
    icon: '/icons/excel.svg',
    href: isKo.value ? '/ko/excel/playbook' : '/excel/playbook',
    meta: isKo.value ? '10개 탭 · .xlsx · 40개 함수' : '10 tabs · .xlsx · 40 formulas'
  },
  {
    id: 'word',
    cls: 's2i-pb-toc-word',
    name: 'Word',
    icon: '/icons/word.svg',
    href: isKo.value ? '/ko/word/playbook' : '/word/playbook',
    meta: isKo.value ? '6페이지 · .docx · 라이브 메타데이터' : '6 pages · .docx · live metadata'
  },
  {
    id: 'ppt',
    cls: 's2i-pb-toc-ppt',
    name: 'PowerPoint',
    icon: '/icons/powerpoint.svg',
    href: isKo.value ? '/ko/powerpoint/playbook' : '/powerpoint/playbook',
    meta: isKo.value ? '8슬라이드 · .pptx · 다크 모드' : '8 slides · .pptx · dark mode'
  }
])
</script>

<template>
  <div
    class="s2i-pb-toc vp-raw"
    :aria-label="isKo ? 'Office 앱별 Playbook 가이드 이동' : 'Choose an Office Playbook guide'"
  >
    <a
      v-for="card in cards"
      :key="card.id"
      class="s2i-pb-toc-card"
      :class="card.cls"
      :href="withBase(card.href)"
    >
      <span class="s2i-pb-toc-head">
        <img :src="withBase(card.icon)" alt="" width="18" height="18" aria-hidden="true" />
        <span>{{ card.name }}</span>
        <span class="s2i-pb-toc-arrow" aria-hidden="true">→</span>
      </span>
      <span class="s2i-pb-toc-meta">{{ card.meta }}</span>
    </a>
  </div>
</template>

<style scoped>
.s2i-pb-toc {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 12px;
  margin: 20px 0 24px;
}

.s2i-pb-toc-card {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 13px 15px;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  background-color: var(--vp-c-bg-soft);
  text-decoration: none;
  transition: border-color 0.2s ease, background-color 0.2s ease, transform 0.2s ease;
}

.s2i-pb-toc-card:hover {
  text-decoration: none;
  transform: translateY(-1px);
}

.s2i-pb-toc-head {
  display: flex;
  align-items: center;
  gap: 8px;
  font-weight: 600;
  font-size: 15px;
  line-height: 1.3;
}

.s2i-pb-toc-head img {
  width: 18px;
  height: 18px;
  flex-shrink: 0;
  display: inline-block;
  margin: 0;
}

.s2i-pb-toc-arrow {
  margin-left: auto;
  font-size: 14px;
  color: var(--vp-c-text-3);
  transition: transform 0.2s ease, color 0.2s ease;
}

.s2i-pb-toc-card:hover .s2i-pb-toc-arrow {
  transform: translateX(3px);
}

.s2i-pb-toc-meta {
  font-size: 12.5px;
  line-height: 1.4;
  color: var(--vp-c-text-2);
}

.s2i-pb-toc-excel .s2i-pb-toc-head {
  color: var(--s2i-excel);
}
.s2i-pb-toc-excel:hover {
  border-color: var(--s2i-excel-border);
  background-color: var(--s2i-excel-bg);
}
.s2i-pb-toc-excel:hover .s2i-pb-toc-arrow {
  color: var(--s2i-excel);
}

.s2i-pb-toc-word .s2i-pb-toc-head {
  color: var(--s2i-word);
}
.s2i-pb-toc-word:hover {
  border-color: var(--s2i-word-border);
  background-color: var(--s2i-word-bg);
}
.s2i-pb-toc-word:hover .s2i-pb-toc-arrow {
  color: var(--s2i-word);
}

.s2i-pb-toc-ppt .s2i-pb-toc-head {
  color: var(--s2i-ppt);
}
.s2i-pb-toc-ppt:hover {
  border-color: var(--s2i-ppt-border);
  background-color: var(--s2i-ppt-bg);
}
.s2i-pb-toc-ppt:hover .s2i-pb-toc-arrow {
  color: var(--s2i-ppt);
}

@media (max-width: 720px) {
  .s2i-pb-toc {
    grid-template-columns: 1fr;
  }
}
</style>
