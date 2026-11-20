<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'
import manifestData from '../mediaManifest.json'

interface MediaEntry {
  id: string
  titleEn: string
  titleKo: string
  shot: string
  webm: string
  mp4: string
  captionEn: string
  captionKo: string
  vttEn: string
  vttKo: string
}

const props = withDefaults(
  defineProps<{
    id: string
    initialTab?: 'video' | 'screenshot'
    titleEn?: string
    titleKo?: string
    captionEn?: string
    captionKo?: string
  }>(),
  {
    initialTab: 'video'
  }
)

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const mode = ref<'video' | 'screenshot'>(props.initialTab)
const zoomOpen = ref(false)

const manifest = manifestData as Record<string, MediaEntry>
const entry = computed<MediaEntry | null>(() => manifest[props.id] ?? null)

const displayTitle = computed(() => {
  if (isKo.value) return props.titleKo ?? entry.value?.titleKo ?? props.id
  return props.titleEn ?? entry.value?.titleEn ?? props.id
})

const displayCaption = computed(() => {
  if (isKo.value) return props.captionKo ?? entry.value?.captionKo ?? displayTitle.value
  return props.captionEn ?? entry.value?.captionEn ?? displayTitle.value
})
</script>

<template>
  <figure v-if="entry" class="s2i-figure">
    <video
      v-if="mode === 'video'"
      class="s2i-media"
      controls
      autoplay
      muted
      loop
      playsinline
      preload="metadata"
      :poster="withBase(entry.shot)"
    >
      <source :src="withBase(entry.webm)" type="video/webm" />
      <source :src="withBase(entry.mp4)" type="video/mp4" />
      <track kind="captions" srclang="ko" label="한국어" :src="withBase(entry.vttKo)" :default="isKo" />
      <track kind="captions" srclang="en" label="English" :src="withBase(entry.vttEn)" :default="!isKo" />
    </video>

    <button
      v-else
      type="button"
      class="s2i-img-btn"
      :title="isKo ? '클릭해서 확대' : 'Click to enlarge'"
      @click="zoomOpen = true"
    >
      <img :src="withBase(entry.shot)" :alt="displayTitle" loading="lazy" class="s2i-media" />
    </button>

    <figcaption class="s2i-figcaption">
      <span class="s2i-caption-text">{{ displayCaption }}</span>
      <span class="s2i-mode-switch" role="tablist">
        <button type="button" role="tab" :aria-selected="mode === 'video'" :class="{ active: mode === 'video' }" @click="mode = 'video'">
          {{ isKo ? '영상' : 'Video' }}
        </button>
        <span class="s2i-sep" aria-hidden="true">/</span>
        <button type="button" role="tab" :aria-selected="mode === 'screenshot'" :class="{ active: mode === 'screenshot' }" @click="mode = 'screenshot'">
          {{ isKo ? '이미지' : 'Image' }}
        </button>
      </span>
    </figcaption>

    <div v-if="zoomOpen" class="s2i-lightbox" @click="zoomOpen = false">
      <img :src="withBase(entry.shot)" :alt="displayTitle" />
    </div>
  </figure>
</template>

<style scoped>
.s2i-figure {
  margin: 16px 0 20px;
}

.s2i-media {
  display: block;
  width: 100%;
  height: auto;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.s2i-img-btn {
  display: block;
  width: 100%;
  padding: 0;
  border: none;
  background: none;
  cursor: zoom-in;
}

.s2i-figcaption {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 12px;
  margin-top: 8px;
  font-size: 0.82rem;
  line-height: 1.5;
  color: var(--vp-c-text-2);
}

.s2i-mode-switch {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
}

.s2i-mode-switch button {
  padding: 0;
  border: none;
  background: none;
  font-size: 0.8rem;
  color: var(--vp-c-text-3);
  cursor: pointer;
  transition: color 0.15s;
}

.s2i-mode-switch button:hover,
.s2i-mode-switch button.active {
  color: var(--vp-c-brand-1);
  font-weight: 600;
}

.s2i-sep {
  color: var(--vp-c-divider);
}

.s2i-lightbox {
  position: fixed;
  inset: 0;
  z-index: 9999;
  background: rgba(0, 0, 0, 0.82);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  cursor: zoom-out;
}

.s2i-lightbox img {
  max-width: 94vw;
  max-height: 88vh;
  border-radius: 8px;
}
</style>
