<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'
import manifestData from '../mediaManifest.json'

interface MediaEntry {
  id: string
  host: string
  titleEn: string
  titleKo: string
  shot: string
  webm: string
  mp4: string
  captionEn: string
  captionKo: string
  vttEn: string
  vttKo: string
  version: string
  capturedAt: string
  width: number
  height: number
  durationSec: number
}

const props = withDefaults(
  defineProps<{
    id: string
    initialTab?: 'video' | 'screenshot' | 'split'
    titleEn?: string
    titleKo?: string
    captionEn?: string
    captionKo?: string
  }>(),
  {
    initialTab: 'split'
  }
)

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const mode = ref<'split' | 'video' | 'screenshot'>(props.initialTab)
const zoomOpen = ref(false)

const manifest = manifestData as Record<string, MediaEntry>
const entry = computed<MediaEntry | null>(() => manifest[props.id] ?? null)

const displayTitle = computed(() => {
  if (isKo.value) return props.titleKo ?? entry.value?.titleKo ?? props.id
  return props.titleEn ?? entry.value?.titleEn ?? props.id
})

const displayCaption = computed(() => {
  if (isKo.value) return props.captionKo ?? entry.value?.captionKo ?? ''
  return props.captionEn ?? entry.value?.captionEn ?? ''
})
</script>

<template>
  <div v-if="entry" class="s2i-fm-card">
    <div class="s2i-fm-header">
      <div class="s2i-fm-title-wrap">
        <span class="s2i-fm-live-dot" aria-hidden="true"></span>
        <strong class="s2i-fm-title">{{ displayTitle }}</strong>
      </div>
      <div class="s2i-fm-controls">
        <div class="s2i-fm-tabs" role="tablist">
          <button
            type="button"
            role="tab"
            :aria-selected="mode === 'split'"
            :class="{ active: mode === 'split' }"
            @click="mode = 'split'"
          >
            {{ isKo ? '📸+🎬 나란히 보기' : '📸+🎬 Side-by-Side' }}
          </button>
          <button
            type="button"
            role="tab"
            :aria-selected="mode === 'video'"
            :class="{ active: mode === 'video' }"
            @click="mode = 'video'"
          >
            {{ isKo ? '🎬 사용 영상' : '🎬 Live Video' }}
          </button>
          <button
            type="button"
            role="tab"
            :aria-selected="mode === 'screenshot'"
            :class="{ active: mode === 'screenshot' }"
            @click="mode = 'screenshot'"
          >
            {{ isKo ? '📸 실제 스크린샷' : '📸 Screenshot' }}
          </button>
        </div>
      </div>
    </div>

    <div class="s2i-fm-stage" :class="`mode-${mode}`">
      <div v-if="mode === 'split' || mode === 'screenshot'" class="s2i-fm-pane">
        <div class="s2i-fm-pane-label">
          {{ isKo ? '📸 실제 실행 스크린샷' : '📸 Live Screenshot' }}
        </div>
        <button
          type="button"
          class="s2i-fm-img-btn"
          :title="isKo ? '클릭해서 크게 보기' : 'Click to enlarge'"
          @click="zoomOpen = true"
        >
          <img
            :src="withBase(entry.shot)"
            :alt="displayTitle"
            loading="lazy"
            class="s2i-fm-img"
          />
        </button>
      </div>

      <div v-if="mode === 'split' || mode === 'video'" class="s2i-fm-pane">
        <div class="s2i-fm-pane-label">
          {{ isKo ? `🎬 실제 사용 동영상 (${entry.durationSec}s)` : `🎬 Live Usage Video (${entry.durationSec}s)` }}
        </div>
        <video
          class="s2i-fm-video"
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
          <track
            kind="captions"
            srclang="ko"
            label="한국어"
            :src="withBase(entry.vttKo)"
            :default="isKo"
          />
          <track
            kind="captions"
            srclang="en"
            label="English"
            :src="withBase(entry.vttEn)"
            :default="!isKo"
          />
        </video>
      </div>
    </div>

    <div class="s2i-fm-footer">
      <span v-if="displayCaption" class="s2i-fm-caption">{{ displayCaption }}</span>
      <span class="s2i-fm-meta">
        {{
          isKo
            ? `실제 Windows Office 구동 기록 · v${entry.version} (${entry.width}×${entry.height})`
            : `Live Windows Office capture · v${entry.version} (${entry.width}×${entry.height})`
        }}
      </span>
    </div>

    <div v-if="zoomOpen" class="s2i-fm-lightbox" @click="zoomOpen = false">
      <img :src="withBase(entry.shot)" :alt="displayTitle" />
      <span class="s2i-fm-lightbox-hint">
        {{ isKo ? '아무 곳이나 클릭해 닫기' : 'Click anywhere to close' }}
      </span>
    </div>
  </div>
</template>

<style scoped>
.s2i-fm-card {
  margin: 18px 0 24px 0;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: var(--vp-c-bg-soft);
  overflow: hidden;
  box-shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
}
.s2i-fm-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
  padding: 10px 14px;
  border-bottom: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
}
.s2i-fm-title-wrap {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}
.s2i-fm-live-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: #10b981;
  box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.2);
}
.s2i-fm-title {
  font-size: 0.9rem;
  color: var(--vp-c-text-1);
}
.s2i-fm-tabs {
  display: inline-flex;
  gap: 4px;
  background: var(--vp-c-bg-soft);
  padding: 3px;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
}
.s2i-fm-tabs button {
  padding: 4px 10px;
  border-radius: 6px;
  font-size: 0.76rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
  background: transparent;
  border: none;
  cursor: pointer;
  transition: all 0.15s;
}
.s2i-fm-tabs button.active {
  background: var(--vp-c-brand-1);
  color: #ffffff;
}
.dark .s2i-fm-tabs button.active {
  color: #0f172a;
}
.s2i-fm-stage {
  padding: 12px;
  display: grid;
  gap: 12px;
}
.s2i-fm-stage.mode-split {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}
.s2i-fm-stage.mode-video,
.s2i-fm-stage.mode-screenshot {
  grid-template-columns: 1fr;
}
@media (max-width: 768px) {
  .s2i-fm-stage.mode-split {
    grid-template-columns: 1fr;
  }
}
.s2i-fm-pane {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.s2i-fm-pane-label {
  font-size: 0.76rem;
  font-weight: 700;
  color: var(--vp-c-text-2);
}
.s2i-fm-img-btn {
  display: block;
  width: 100%;
  padding: 0;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  overflow: hidden;
  background: #ffffff;
  cursor: zoom-in;
}
.s2i-fm-img {
  display: block;
  width: 100%;
  height: auto;
  aspect-ratio: 16 / 9;
  object-fit: contain;
  background: #0f172a;
}
.s2i-fm-video {
  display: block;
  width: 100%;
  height: auto;
  aspect-ratio: 16 / 9;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  background: #0f172a;
}
.s2i-fm-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 6px;
  padding: 8px 14px;
  border-top: 1px solid var(--vp-c-divider);
  font-size: 0.78rem;
  color: var(--vp-c-text-2);
}
.s2i-fm-meta {
  font-family: var(--vp-font-family-mono);
  font-size: 0.72rem;
  color: var(--vp-c-text-3);
}
.s2i-fm-lightbox {
  position: fixed;
  inset: 0;
  z-index: 9999;
  background: rgba(15, 23, 42, 0.88);
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  padding: 24px;
  cursor: zoom-out;
}
.s2i-fm-lightbox img {
  max-width: 94vw;
  max-height: 86vh;
  border-radius: 10px;
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
}
.s2i-fm-lightbox-hint {
  color: #e2e8f0;
  font-size: 0.85rem;
  font-weight: 600;
}
</style>
