<script setup lang="ts">
import { ref, computed, watch } from 'vue'
import { useData, withBase } from 'vitepress'
import { openLightbox } from '../lightbox'

type AppId = 'excel' | 'word' | 'ppt'

const props = defineProps<{
  app?: AppId
}>()

interface SlideItem {
  img: string
  captionEn: string
  captionKo: string
}

interface PlaybookTab {
  id: AppId
  name: string
  shortName: string
  icon: string
  ext: string
  file: string
  filename: string
  docEn: string
  docKo: string
  slides: SlideItem[]
}

const tabs: PlaybookTab[] = [
  {
    id: 'excel',
    name: 'Excel Playbook',
    shortName: 'Excel',
    icon: '/icons/excel.svg',
    ext: '.xlsx',
    file: '/examples/func_SMILES2IMG_Excel_Playbook.xlsx',
    filename: 'func_SMILES2IMG_Excel_Playbook.xlsx',
    docEn: '/excel/playbook',
    docKo: '/ko/excel/playbook',
    slides: [
      {
        img: '/examples/previews/excel-01.webp',
        captionEn: 'Tab 0 · Start Here — Interactive sandbox & roadmap',
        captionKo: '탭 0 · 시작하기 — 체험 샌드박스 및 학습 로드맵'
      },
      {
        img: '/examples/previews/excel-02.webp',
        captionEn: 'Tab 1 · Structure Rendering — Name, CAS, SMILES & InChI to 2D',
        captionKo: '탭 1 · 구조 렌더링 — 이름·CAS·SMILES·InChI 2D 자동 변환'
      },
      {
        img: '/examples/previews/excel-03.webp',
        captionEn: 'Tab 2 · Properties & 3D — 15-col physicochemical & 3D KPIs',
        captionKo: '탭 2 · 분자 물성 & 3D — 15열 물성 및 3D 복잡도 지표'
      },
      {
        img: '/examples/previews/excel-04.webp',
        captionEn: 'Tab 3 · SAR & SMARTS — SMARTS alerts, hl: highlights & InChI parity',
        captionKo: '탭 3 · SAR & SMARTS — 부분구조 하이라이트 및 InChI 동등 검색'
      },
      {
        img: '/examples/previews/excel-05.webp',
        captionEn: 'Tab 4 · Notation Conversions — InChI/InChIKey & deduplication',
        captionKo: '탭 4 · 표기법 변환 — InChI·InChIKey 표준화 및 중복 검출'
      },
      {
        img: '/examples/previews/excel-06.webp',
        captionEn: 'Tab 5 · Styles & Transforms — ACS B&W, dark mode & 16 presets',
        captionKo: '탭 5 · 스타일 & 변환 — 학술지 흑백·다크모드 등 16종 프리셋'
      },
      {
        img: '/examples/previews/excel-07.webp',
        captionEn: 'Tab 6 · Reactions & KET — Multi-step synthesis schemes & Ketcher',
        captionKo: '탭 6 · 반응식 & KET — 다단계 합성 반응식 및 무손실 KET 편집'
      },
      {
        img: '/examples/previews/excel-08.webp',
        captionEn: 'Tab 7 · Real-World Workflows — Stoichiometry & hit-to-lead triage',
        captionKo: '탭 7 · 실전 워크플로우 — 화학양론 계산기 및 Hit-to-Lead 선별표'
      },
      {
        img: '/examples/previews/excel-09.webp',
        captionEn: 'Tab 8 · Legacy Excel (.FLOAT) — Excel 2016–2021 floating shapes',
        captionKo: '탭 8 · 구버전 Excel (.FLOAT) — Excel 2016–2021 호환 부동 도형'
      },
      {
        img: '/examples/previews/excel-10.webp',
        captionEn: 'Tab 9 · Ribbon, SDF & Ref — Complete 49-function offline reference',
        captionKo: '탭 9 · 리본 · SDF · 사전 — 49종 전체 함수 및 오프라인 해결 가이드'
      }
    ]
  },
  {
    id: 'word',
    name: 'Word Playbook',
    shortName: 'Word',
    icon: '/icons/word.svg',
    ext: '.docx',
    file: '/examples/func_SMILES2IMG_Word_Playbook.docx',
    filename: 'func_SMILES2IMG_Word_Playbook.docx',
    docEn: '/word/playbook',
    docKo: '/ko/word/playbook',
    slides: [
      {
        img: '/examples/previews/word-01.webp',
        captionEn: 'Page 1 of 6 · Cover Banner, Quick Guide & [SMILES] Ribbon Reference Table',
        captionKo: '1 / 6쪽 · 표지 배너, 실습 안내 및 [SMILES] 리본·우클릭 메뉴 요약표'
      },
      {
        img: '/examples/previews/word-02.webp',
        captionEn: 'Page 2 of 6 · Ch. 1 Inline 2D Structures & Text-to-Structure Conversion Lab',
        captionKo: '2 / 6쪽 · 1장 본문 인라인 2D 구조식 및 텍스트 제자리 변환 실습'
      },
      {
        img: '/examples/previews/word-03.webp',
        captionEn: 'Page 3 of 6 · Ch. 2 Journal Compound Summary & SAR Table + Molecule Info',
        captionKo: '3 / 6쪽 · 2장 학술지 화합물 요약·SAR 표 및 Molecule Info 물성 추출'
      },
      {
        img: '/examples/previews/word-04.webp',
        captionEn: 'Page 4 of 6 · Ch. 3 Multi-Step Organic Reaction Schemes (Lossless KET)',
        captionKo: '4 / 6쪽 · 3장 다단계 유기 합성 반응식 (무손실 KET 메타데이터)'
      },
      {
        img: '/examples/previews/word-05.webp',
        captionEn: 'Page 5 of 6 · Ch. 4 Mol* 3D Macromolecular Target Figure (PDB: 1CRN)',
        captionKo: '5 / 6쪽 · 4장 Mol* 3D 단백질 구조 Figure (PDB: 1CRN Crambin)'
      },
      {
        img: '/examples/previews/word-06.webp',
        captionEn: 'Page 6 of 6 · Ch. 5 ACS B&W Batch Style, Vector SVG & Troubleshooting',
        captionKo: '6 / 6쪽 · 5장 학술지 흑백 일괄 변환(Batch Style), 벡터 SVG 및 문제 해결'
      }
    ]
  },
  {
    id: 'ppt',
    name: 'PowerPoint Playbook',
    shortName: 'PowerPoint',
    icon: '/icons/powerpoint.svg',
    ext: '.pptx',
    file: '/examples/func_SMILES2IMG_PowerPoint_Playbook.pptx',
    filename: 'func_SMILES2IMG_PowerPoint_Playbook.pptx',
    docEn: '/powerpoint/playbook',
    docKo: '/ko/powerpoint/playbook',
    slides: [
      {
        img: '/examples/previews/ppt-01.webp',
        captionEn: 'Slide 1 · Dark-Mode Conference Deck Cover & Hands-On Roadmap',
        captionKo: '슬라이드 1 · 다크모드 학회 발표 표지 및 실습 구성 안내'
      },
      {
        img: '/examples/previews/ppt-02.webp',
        captionEn: 'Slide 2 · Complete PowerPoint [SMILES] Ribbon & Right-Click Menu Map',
        captionKo: '슬라이드 2 · PowerPoint [SMILES] 리본 및 3단계 우클릭 메뉴 요약표'
      },
      {
        img: '/examples/previews/ppt-03.webp',
        captionEn: 'Slide 3 · TextBox Conversion Lab (SMILES / Name / Chiral)',
        captionKo: '슬라이드 3 · 텍스트 상자 구조 변환 실습 (SMILES · 이름 · 카이랄)'
      },
      {
        img: '/examples/previews/ppt-04.webp',
        captionEn: 'Slide 4 · 4-Card Lead Screening Grid with Dark-Mode Structures',
        captionKo: '슬라이드 4 · 투명 다크모드 4-카드 후보물질 스크리닝 레이아웃'
      },
      {
        img: '/examples/previews/ppt-05.webp',
        captionEn: 'Slide 5 · Double-Click Editable KET Reaction Mechanism Slide',
        captionKo: '슬라이드 5 · 더블클릭 편집 가능한 KET 반응 메커니즘 슬라이드'
      },
      {
        img: '/examples/previews/ppt-06.webp',
        captionEn: 'Slide 6 · Embedded Mol* 3D Protein Target (1CRN) & .glb Model',
        captionKo: '슬라이드 6 · 임베드 Mol* 3D 단백질 표적(1CRN) 및 .glb 회전 모델'
      },
      {
        img: '/examples/previews/ppt-07.webp',
        captionEn: 'Slide 7 · Q&A Molecule Info & Batch Style Switching',
        captionKo: '슬라이드 7 · 발표 Q&A용 분자 물성 조회 및 스타일 일괄 변환 실습'
      },
      {
        img: '/examples/previews/ppt-08.webp',
        captionEn: 'Slide 8 · Sharing Decks, A0 Poster Vector SVG & Troubleshooting',
        captionKo: '슬라이드 8 · macOS/발표장 공유, A0 포스터 벡터 SVG 및 문제 해결'
      }
    ]
  }
]

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const activeId = ref<AppId>(props.app ?? 'excel')
const slideIndex = ref(0)

watch(
  () => props.app,
  (next) => {
    if (next) {
      activeId.value = next
      slideIndex.value = 0
    }
  }
)

const activeTab = computed(() => tabs.find((t) => t.id === (props.app ?? activeId.value)) ?? tabs[0])
const activeSlide = computed(() => activeTab.value.slides[slideIndex.value] ?? activeTab.value.slides[0])
const totalSlides = computed(() => activeTab.value.slides.length)

function selectTab(id: AppId) {
  if (activeId.value === id) return
  activeId.value = id
  slideIndex.value = 0
}

function prevSlide() {
  slideIndex.value = (slideIndex.value - 1 + totalSlides.value) % totalSlides.value
}

function nextSlide() {
  slideIndex.value = (slideIndex.value + 1) % totalSlides.value
}

function goToSlide(idx: number) {
  slideIndex.value = idx
}

function onStageKeydown(e: KeyboardEvent) {
  if (e.key === 'ArrowLeft') {
    e.preventDefault()
    prevSlide()
  } else if (e.key === 'ArrowRight') {
    e.preventDefault()
    nextSlide()
  }
}

function openSlideLightbox() {
  const slides = activeTab.value.slides.map((s) => {
    const caption = isKo.value ? s.captionKo : s.captionEn
    return {
      src: withBase(s.img),
      alt: caption,
      caption
    }
  })
  openLightbox(slides, slideIndex.value, (idx) => {
    slideIndex.value = idx
  })
}
</script>

<template>
  <div class="s2i-pb vp-raw" :class="[`is-${activeTab.id}`, { 'is-single': Boolean(props.app) }]">
    <div v-if="!props.app" class="s2i-pb-tabs" role="tablist" :aria-label="isKo ? 'Playbook 앱 선택' : 'Select Playbook app'">
      <button
        v-for="tab in tabs"
        :key="tab.id"
        type="button"
        role="tab"
        class="s2i-pb-tab"
        :class="[`tab-${tab.id}`, { active: activeId === tab.id }]"
        :aria-selected="activeId === tab.id"
        @click="selectTab(tab.id)"
      >
        <img :src="withBase(tab.icon)" alt="" width="20" height="20" aria-hidden="true" />
        <span class="s2i-pb-tab-title">{{ tab.shortName }}</span>
      </button>
    </div>

    <div
      class="s2i-pb-stage"
      :role="props.app ? 'region' : 'tabpanel'"
      tabindex="0"
      :aria-label="activeTab.name"
      @keydown="onStageKeydown"
    >
      <button
        type="button"
        class="s2i-pb-img-wrap"
        :title="isKo ? '클릭해서 크게 보기' : 'Click to enlarge'"
        @click="openSlideLightbox"
      >
        <img
          :key="activeSlide.img"
          :src="withBase(activeSlide.img)"
          :alt="isKo ? activeSlide.captionKo : activeSlide.captionEn"
          class="s2i-pb-img"
        />
      </button>

      <button
        type="button"
        class="s2i-pb-nav s2i-pb-prev"
        :aria-label="isKo ? '이전 미리보기' : 'Previous preview'"
        @click="prevSlide"
      >
        <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="m15 18-6-6 6-6" />
        </svg>
      </button>

      <button
        type="button"
        class="s2i-pb-nav s2i-pb-next"
        :aria-label="isKo ? '다음 미리보기' : 'Next preview'"
        @click="nextSlide"
      >
        <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="m9 18 6-6-6-6" />
        </svg>
      </button>
    </div>

    <div class="s2i-pb-footer">
      <div class="s2i-pb-info">
        <div class="s2i-pb-caption-row">
          <span class="s2i-pb-counter">{{ slideIndex + 1 }} / {{ totalSlides }}</span>
          <span class="s2i-pb-caption">{{ isKo ? activeSlide.captionKo : activeSlide.captionEn }}</span>
        </div>
        <div class="s2i-pb-dots" role="tablist" :aria-label="isKo ? '슬라이드 선택' : 'Select preview slide'">
          <button
            v-for="(slide, idx) in activeTab.slides"
            :key="slide.img"
            type="button"
            class="s2i-pb-dot"
            :class="{ active: idx === slideIndex }"
            :title="isKo ? slide.captionKo : slide.captionEn"
            :aria-label="`${idx + 1} / ${totalSlides}`"
            :aria-current="idx === slideIndex ? 'true' : undefined"
            @click="goToSlide(idx)"
          />
        </div>
      </div>

      <div class="s2i-pb-actions">
        <a
          v-if="!props.app"
          class="s2i-pb-dl-btn"
          :href="withBase(isKo ? activeTab.docKo : activeTab.docEn)"
        >
          <span>{{ isKo ? `${activeTab.name} 보기` : activeTab.name }}</span>
          <svg class="s2i-pb-btn-arrow" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M5 12h14" />
            <path d="m12 5 7 7-7 7" />
          </svg>
        </a>
        <a
          v-else
          class="s2i-pb-dl-btn"
          :href="withBase(activeTab.file)"
          :download="activeTab.filename"
        >
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" />
            <path d="M14 2v4a2 2 0 0 0 2 2h4" />
            <path d="M12 18v-6" />
            <path d="m9 15 3 3 3-3" />
          </svg>
          <span>{{ isKo ? `${activeTab.ext} 다운로드` : `Download ${activeTab.ext}` }}</span>
        </a>
      </div>
    </div>
  </div>
</template>

<style scoped>
.s2i-pb {
  margin-top: 16px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: var(--vp-c-bg-soft);
  overflow: hidden;
}

/* 3 Equal-Width Top Tabs */
.s2i-pb-tabs {
  display: flex;
  border-bottom: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.s2i-pb-tab {
  position: relative;
  flex: 1 1 0;
  min-width: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 9px;
  padding: 13px 16px;
  border: none;
  border-right: 1px solid var(--vp-c-divider);
  background: transparent;
  cursor: pointer;
  color: var(--vp-c-text-2);
  transition: background-color 0.18s ease, color 0.18s ease;
}

.s2i-pb-tab:last-child {
  border-right: none;
}

.s2i-pb-tab::after {
  content: '';
  position: absolute;
  left: 0;
  right: 0;
  bottom: -1px;
  height: 2px;
  background: transparent;
  transition: background-color 0.18s ease;
}

.s2i-pb-tab img {
  width: 20px;
  height: 20px;
  flex-shrink: 0;
  opacity: 0.78;
  transition: opacity 0.18s ease;
}

.s2i-pb-tab:hover img,
.s2i-pb-tab.active img {
  opacity: 1;
}

.s2i-pb-tab-title {
  font-size: 14.5px;
  font-weight: 600;
  color: var(--vp-c-text-1);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.s2i-pb-tab:hover {
  background: var(--vp-c-bg-elv);
}

.s2i-pb-tab.tab-excel.active {
  background: var(--s2i-excel-bg);
}
.s2i-pb-tab.tab-excel.active .s2i-pb-tab-title {
  color: var(--s2i-excel);
}
.s2i-pb-tab.tab-excel.active::after {
  background: var(--s2i-excel-border);
}

.s2i-pb-tab.tab-word.active {
  background: var(--s2i-word-bg);
}
.s2i-pb-tab.tab-word.active .s2i-pb-tab-title {
  color: var(--s2i-word);
}
.s2i-pb-tab.tab-word.active::after {
  background: var(--s2i-word-border);
}

.s2i-pb-tab.tab-ppt.active {
  background: var(--s2i-ppt-bg);
}
.s2i-pb-tab.tab-ppt.active .s2i-pb-tab-title {
  color: var(--s2i-ppt);
}
.s2i-pb-tab.tab-ppt.active::after {
  background: var(--s2i-ppt-border);
}

/* 3:2 Carousel Stage */
.s2i-pb-stage {
  position: relative;
  width: 100%;
  aspect-ratio: 3 / 2;
  background: #141820;
  outline: none;
}

.s2i-pb-img-wrap {
  display: block;
  width: 100%;
  height: 100%;
  padding: 0;
  border: none;
  background: transparent;
  cursor: zoom-in;
}

.s2i-pb-img {
  width: 100%;
  height: 100%;
  object-fit: contain;
  display: block;
  user-select: none;
}

.s2i-pb-nav {
  position: absolute;
  top: 50%;
  transform: translateY(-50%);
  width: 38px;
  height: 38px;
  border-radius: 50%;
  border: 1px solid rgba(255, 255, 255, 0.22);
  background: rgba(15, 23, 42, 0.62);
  color: #f8fafc;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  backdrop-filter: blur(4px);
  transition: background-color 0.16s ease, border-color 0.16s ease, transform 0.16s ease;
}

.s2i-pb-nav:hover {
  background: rgba(15, 23, 42, 0.88);
  border-color: rgba(255, 255, 255, 0.45);
}

.s2i-pb-prev {
  left: 14px;
}

.s2i-pb-next {
  right: 14px;
}

/* Bottom Footer Bar: Caption + Dots (left) | Action CTA (right) */
.s2i-pb-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 13px 18px;
  border-top: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.s2i-pb-info {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-width: 0;
  flex: 1;
}

.s2i-pb-caption-row {
  display: flex;
  align-items: baseline;
  gap: 9px;
  min-width: 0;
}

.s2i-pb-counter {
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
  font-weight: 600;
  color: var(--vp-c-text-3);
  flex-shrink: 0;
}

.s2i-pb-caption {
  font-size: 13.5px;
  line-height: 1.4;
  color: var(--vp-c-text-1);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.s2i-pb-dots {
  display: flex;
  align-items: center;
  gap: 6px;
}

.s2i-pb-dot {
  width: 7px;
  height: 7px;
  padding: 0;
  border: none;
  border-radius: 999px;
  background: var(--vp-c-divider);
  cursor: pointer;
  transition: width 0.2s ease, background-color 0.2s ease;
}

.s2i-pb-dot:hover {
  background: var(--vp-c-text-3);
}

.s2i-pb-dot.active {
  width: 20px;
}

.s2i-pb.is-excel .s2i-pb-dot.active {
  background: var(--s2i-excel-border);
}
.s2i-pb.is-word .s2i-pb-dot.active {
  background: var(--s2i-word-border);
}
.s2i-pb.is-ppt .s2i-pb-dot.active {
  background: var(--s2i-ppt-border);
}

.s2i-pb-actions {
  display: flex;
  align-items: center;
  flex-shrink: 0;
}

.s2i-pb-dl-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 7px;
  padding: 8px 15px;
  border-radius: 8px;
  font-size: 13.5px;
  font-weight: 600;
  text-decoration: none;
  white-space: nowrap;
  border: 1px solid transparent;
  transition: background-color 0.16s ease, border-color 0.16s ease, filter 0.16s ease;
}

.s2i-pb-btn-arrow {
  transition: transform 0.16s ease;
}

.s2i-pb-dl-btn:hover .s2i-pb-btn-arrow {
  transform: translateX(2px);
}

.s2i-pb.is-excel .s2i-pb-dl-btn {
  color: var(--s2i-excel);
  background: var(--s2i-excel-bg);
  border-color: var(--s2i-excel-border);
}

.s2i-pb.is-word .s2i-pb-dl-btn {
  color: var(--s2i-word);
  background: var(--s2i-word-bg);
  border-color: var(--s2i-word-border);
}

.s2i-pb.is-ppt .s2i-pb-dl-btn {
  color: var(--s2i-ppt);
  background: var(--s2i-ppt-bg);
  border-color: var(--s2i-ppt-border);
}

.s2i-pb-dl-btn:hover {
  filter: brightness(1.08);
}

@media (max-width: 767px) {
  .s2i-pb-tab {
    padding: 11px 8px;
    gap: 6px;
  }

  .s2i-pb-tab img {
    width: 18px;
    height: 18px;
  }

  .s2i-pb-tab-title {
    font-size: 13px;
  }

  .s2i-pb-nav {
    width: 32px;
    height: 32px;
  }

  .s2i-pb-prev {
    left: 8px;
  }

  .s2i-pb-next {
    right: 8px;
  }

  .s2i-pb-footer {
    flex-direction: column;
    align-items: stretch;
    gap: 12px;
    padding: 12px 14px;
  }

  .s2i-pb-caption {
    white-space: normal;
  }

  .s2i-pb-actions {
    justify-content: flex-end;
  }
}
</style>
