<script setup lang="ts">
import { computed, watch, onMounted, onUnmounted } from 'vue'
import { useRoute, useData } from 'vitepress'
import {
  lightboxState,
  openLightbox,
  closeLightbox,
  prevLightboxSlide,
  nextLightboxSlide
} from '../lightbox'

const route = useRoute()
const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const current = computed(() => lightboxState.slides[lightboxState.index] ?? null)
const hasMultiple = computed(() => lightboxState.slides.length > 1)

function isZoomableImage(img: HTMLImageElement): boolean {
  if (img.getAttribute('aria-hidden') === 'true') return false
  if (img.classList.contains('s2i-no-zoom')) return false
  if (img.closest('.s2i-pb, .s2i-figure, .s2i-app-label, .s2i-lb-overlay, a, button')) return false
  if (img.src.includes('/icons/')) return false

  const inDoc = Boolean(img.closest('.vp-doc'))
  const isHero = img.classList.contains('image-src') && Boolean(img.closest('.VPHomeHero'))
  if (!inDoc && !isHero) return false

  const rect = img.getBoundingClientRect()
  if (rect.width < 80 && rect.height < 80) return false

  return Boolean(img.currentSrc || img.src)
}

function onDocumentClick(e: MouseEvent) {
  if (lightboxState.open) return
  const target = e.target
  if (!(target instanceof HTMLImageElement)) return
  if (!isZoomableImage(target)) return

  e.preventDefault()
  const alt = target.alt || ''
  openLightbox([{ src: target.currentSrc || target.src, alt, caption: alt }])
}

function onKeydown(e: KeyboardEvent) {
  if (!lightboxState.open) return
  if (e.key === 'Escape') {
    closeLightbox()
  } else if (e.key === 'ArrowLeft' && hasMultiple.value) {
    e.preventDefault()
    prevLightboxSlide()
  } else if (e.key === 'ArrowRight' && hasMultiple.value) {
    e.preventDefault()
    nextLightboxSlide()
  }
}

watch(
  () => route.path,
  () => {
    if (lightboxState.open) closeLightbox()
  }
)

watch(
  () => lightboxState.open,
  (isOpen) => {
    if (typeof document === 'undefined') return
    document.body.style.overflow = isOpen ? 'hidden' : ''
  }
)

onMounted(() => {
  document.addEventListener('click', onDocumentClick)
  window.addEventListener('keydown', onKeydown)
})

onUnmounted(() => {
  document.removeEventListener('click', onDocumentClick)
  window.removeEventListener('keydown', onKeydown)
  if (typeof document !== 'undefined') {
    document.body.style.overflow = ''
  }
})
</script>

<template>
  <Transition name="s2i-lb-fade">
    <div
      v-if="lightboxState.open && current"
      class="s2i-lb-overlay"
      role="dialog"
      aria-modal="true"
      :aria-label="current.alt || 'Enlarged image preview'"
      @click="closeLightbox"
    >
      <button
        type="button"
        class="s2i-lb-btn s2i-lb-close"
        :aria-label="isKo ? '닫기' : 'Close'"
        @click.stop="closeLightbox"
      >
        <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="M18 6 6 18" />
          <path d="m6 6 12 12" />
        </svg>
      </button>

      <button
        v-if="hasMultiple"
        type="button"
        class="s2i-lb-btn s2i-lb-nav s2i-lb-prev"
        :aria-label="isKo ? '이전 이미지' : 'Previous image'"
        @click.stop="prevLightboxSlide"
      >
        <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="m15 18-6-6 6-6" />
        </svg>
      </button>

      <figure class="s2i-lb-figure" @click="closeLightbox">
        <img :key="current.src" :src="current.src" :alt="current.alt" class="s2i-lb-img" />
        <figcaption v-if="current.caption || hasMultiple" class="s2i-lb-caption">
          <span v-if="hasMultiple" class="s2i-lb-counter">
            {{ lightboxState.index + 1 }} / {{ lightboxState.slides.length }}
          </span>
          <span v-if="current.caption">{{ current.caption }}</span>
        </figcaption>
      </figure>

      <button
        v-if="hasMultiple"
        type="button"
        class="s2i-lb-btn s2i-lb-nav s2i-lb-next"
        :aria-label="isKo ? '다음 이미지' : 'Next image'"
        @click.stop="nextLightboxSlide"
      >
        <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="m9 18 6-6-6-6" />
        </svg>
      </button>
    </div>
  </Transition>
</template>

<style scoped>
.s2i-lb-overlay {
  position: fixed;
  inset: 0;
  z-index: 9999;
  background: rgba(0, 0, 0, 0.86);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 28px 64px;
  cursor: zoom-out;
}

.s2i-lb-btn {
  position: fixed;
  width: 40px;
  height: 40px;
  border-radius: 50%;
  border: 1px solid rgba(255, 255, 255, 0.22);
  background: rgba(15, 23, 42, 0.72);
  color: #f8fafc;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  backdrop-filter: blur(4px);
  transition: background-color 0.16s ease, border-color 0.16s ease;
}

.s2i-lb-btn:hover {
  background: rgba(15, 23, 42, 0.95);
  border-color: rgba(255, 255, 255, 0.5);
}

.s2i-lb-close {
  top: 18px;
  right: 20px;
}

.s2i-lb-nav {
  top: 50%;
  transform: translateY(-50%);
}

.s2i-lb-prev {
  left: 18px;
}

.s2i-lb-next {
  right: 18px;
}

.s2i-lb-figure {
  margin: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
  max-width: 92vw;
  max-height: 90vh;
}

.s2i-lb-img {
  max-width: 90vw;
  max-height: 82vh;
  width: auto;
  height: auto;
  border-radius: 8px;
  box-shadow: 0 20px 50px rgba(0, 0, 0, 0.5);
  user-select: none;
}

.s2i-lb-caption {
  display: inline-flex;
  align-items: baseline;
  gap: 10px;
  font-size: 13.5px;
  line-height: 1.45;
  color: rgba(248, 250, 252, 0.88);
  text-align: center;
  max-width: 78ch;
}

.s2i-lb-counter {
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
  font-weight: 600;
  color: rgba(248, 250, 252, 0.62);
  flex-shrink: 0;
}

.s2i-lb-fade-enter-active,
.s2i-lb-fade-leave-active {
  transition: opacity 0.16s ease;
}

.s2i-lb-fade-enter-from,
.s2i-lb-fade-leave-to {
  opacity: 0;
}

@media (max-width: 767px) {
  .s2i-lb-overlay {
    padding: 20px 14px;
  }

  .s2i-lb-btn {
    width: 34px;
    height: 34px;
  }

  .s2i-lb-prev {
    left: 10px;
  }

  .s2i-lb-next {
    right: 10px;
  }

  .s2i-lb-img {
    max-width: 94vw;
  }
}
</style>
