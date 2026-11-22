import { reactive } from 'vue'

export interface LightboxSlide {
  src: string
  alt: string
  caption?: string
}

interface LightboxState {
  open: boolean
  slides: LightboxSlide[]
  index: number
  onChange: ((idx: number) => void) | null
}

export const lightboxState = reactive<LightboxState>({
  open: false,
  slides: [],
  index: 0,
  onChange: null
})

export function openLightbox(
  slides: LightboxSlide[],
  startIndex = 0,
  onChange?: (idx: number) => void
) {
  if (!slides.length) return
  lightboxState.slides = slides
  lightboxState.index = Math.max(0, Math.min(startIndex, slides.length - 1))
  lightboxState.onChange = onChange ?? null
  lightboxState.open = true
}

export function closeLightbox() {
  lightboxState.open = false
  lightboxState.onChange = null
}

export function prevLightboxSlide() {
  const len = lightboxState.slides.length
  if (len <= 1) return
  lightboxState.index = (lightboxState.index - 1 + len) % len
  lightboxState.onChange?.(lightboxState.index)
}

export function nextLightboxSlide() {
  const len = lightboxState.slides.length
  if (len <= 1) return
  lightboxState.index = (lightboxState.index + 1) % len
  lightboxState.onChange?.(lightboxState.index)
}
