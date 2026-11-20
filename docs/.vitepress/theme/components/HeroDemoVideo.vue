<script setup lang="ts">
import { ref, computed, onMounted, onBeforeUnmount } from 'vue'
import { useData, withBase } from 'vitepress'

const { lang } = useData()
const videoEl = ref<HTMLVideoElement | null>(null)
let observer: IntersectionObserver | null = null

const ariaLabel = computed(() =>
  lang.value.startsWith('ko')
    ? 'func_SMILES2IMG — 엑셀 수식 구조식 자동 생성'
    : 'func_SMILES2IMG — Auto-render chemical structures directly in Excel'
)

const togglePlayback = () => {
  const el = videoEl.value
  if (!el) return
  if (el.paused) {
    el.muted = true
    el.play().catch(() => {})
  } else {
    el.pause()
  }
}

onMounted(() => {
  const el = videoEl.value
  if (!el) return

  el.muted = true
  const reducedMotion = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches

  if (reducedMotion) {
    el.pause()
    return
  }

  el.play().catch(() => {})

  if (typeof IntersectionObserver !== 'undefined') {
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            el.muted = true
            el.play().catch(() => {})
          } else {
            el.pause()
          }
        }
      },
      { threshold: 0.1 }
    )
    observer.observe(el)
  }
})

onBeforeUnmount(() => {
  observer?.disconnect()
  observer = null
})
</script>

<template>
  <video
    ref="videoEl"
    class="image-src s2i-hero-video"
    width="1280"
    height="720"
    autoplay
    muted
    loop
    playsinline
    preload="auto"
    :poster="withBase('/hero-demo-poster.webp')"
    :aria-label="ariaLabel"
    @click="togglePlayback"
  >
    <source :src="withBase('/hero-demo.webm')" type="video/webm" />
  </video>
</template>

<style scoped>
.s2i-hero-video {
  display: block;
  aspect-ratio: 16 / 9;
  object-fit: cover;
  background: var(--vp-c-bg-soft);
  cursor: pointer;
}
</style>
