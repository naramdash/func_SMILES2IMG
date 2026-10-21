<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData } from 'vitepress'

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const activePreset = ref<'color' | 'bw' | 'whitebg' | 'dark'>('color')
const surface = ref<'light' | 'shaded' | 'dark'>('light')

function pickPreset(preset: 'color' | 'bw' | 'whitebg' | 'dark') {
  activePreset.value = preset
  if (preset === 'dark') surface.value = 'dark'
  else if (preset === 'whitebg') surface.value = 'shaded'
  else surface.value = 'light'
}

const presetInfo = computed(() => {
  if (isKo.value) {
    switch (activePreset.value) {
      case 'bw':
        return {
          title: 'Publication B&W (ACS 논문·특허용 흑백)',
          desc: '원소 색상을 모두 제거하고 고대비 검은색 선과 글자로 통일합니다. ACS, Wiley, Elsevier 등 인쇄 학술지 투고 및 특허 명세서 제출 시 표준 규격입니다.',
          formula: '=SMILES2IMG(A2, , "bw")'
        }
      case 'whitebg':
        return {
          title: 'White Background (불투명 흰색 사각형 배경)',
          desc: '분자 뒤에 깔끔한 흰색 사각형 배경을 깔아줍니다. 줄무늬 색상이 칠해진 엑셀 표, 음영이 들어간 워드 표, 사진 배경 슬라이드 위에서도 분자가 또렷하게 보입니다.',
          formula: '=SMILES2IMG(A2, "white")'
        }
      case 'dark':
        return {
          title: 'Dark Mode (다크 슬라이드·테마용 흰색 선)',
          desc: '결합선과 탄소 골격을 선명한 흰색 잉크로 바꾸고 투명 배경을 유지합니다. 어두운 학회 발표 슬라이드(PPT)나 다크 모드 엑셀에서 흰 네모 테두리 없이 완벽히 어우러집니다.',
          formula: '=SMILES2IMG(A2, "#1e1e1e", "white")'
        }
      default:
        return {
          title: 'Standard Color (표준 원소 컬러 + 투명 배경)',
          desc: '산소(빨강), 질소(파랑) 등 IUPAC 표준 원소 색상과 투명 배경으로 렌더링합니다. 일반적인 흰색 엑셀 시트, 워드 보고서, 밝은 발표 슬라이드에 가장 적합합니다.',
          formula: '=SMILES2IMG(A2)'
        }
    }
  } else {
    switch (activePreset.value) {
      case 'bw':
        return {
          title: 'Publication B&W (ACS & Patent Monochrome)',
          desc: 'Removes all element colors and renders crisp black bonds and labels. Ideal for print journal manuscripts (ACS, Wiley, Elsevier, RSC) and patent filings.',
          formula: '=SMILES2IMG(A2, , "bw")'
        }
      case 'whitebg':
        return {
          title: 'White Background (Solid White Box)',
          desc: 'Places a solid white background behind the structure so atoms and bonds stay clearly legible over shaded table rows, colored cells, or photo slide backgrounds.',
          formula: '=SMILES2IMG(A2, "white")'
        }
      case 'dark':
        return {
          title: 'Dark Mode (White Ink + Transparent BG)',
          desc: 'Renders bonds in crisp white ink over a transparent background—blending seamlessly into dark PowerPoint slide decks and dark-mode spreadsheets without ugly white boxes.',
          formula: '=SMILES2IMG(A2, "#1e1e1e", "white")'
        }
      default:
        return {
          title: 'Standard Color (IUPAC Color + Transparent BG)',
          desc: 'Uses standard IUPAC element colors (red O, blue N) with a transparent background. Best for everyday spreadsheets, color thesis chapters, and light slide themes.',
          formula: '=SMILES2IMG(A2)'
        }
    }
  }
})

const bondStroke = computed(() => {
  if (activePreset.value === 'dark') return '#f8fafc'
  return '#1e293b'
})

const heteroFill = computed(() => {
  if (activePreset.value === 'bw') return '#1e293b'
  if (activePreset.value === 'dark') return '#f87171'
  return '#dc2626'
})
</script>

<template>
  <div class="s2i-widget bsp-root">
    <div class="s2i-widget-header">
      <div class="s2i-widget-title">
        <span>🎨 {{ isKo ? 'Batch Style 4대 프리셋 시각 비교기 (클릭해서 확인해 보세요)' : 'Batch Style Preset Visualizer (Click to Compare)' }}</span>
      </div>
      <span class="s2i-badge">{{ isKo ? '리본: [SMILES] ➔ [Batch Style]' : 'Ribbon: [SMILES] ➔ [Batch Style]' }}</span>
    </div>

    <div class="bsp-tabs">
      <button type="button" :class="{ active: activePreset === 'color' }" @click="pickPreset('color')">
        🎨 Standard Color
      </button>
      <button type="button" :class="{ active: activePreset === 'bw' }" @click="pickPreset('bw')">
        ✒️ Publication B&amp;W
      </button>
      <button type="button" :class="{ active: activePreset === 'whitebg' }" @click="pickPreset('whitebg')">
        ⬜ White Background
      </button>
      <button type="button" :class="{ active: activePreset === 'dark' }" @click="pickPreset('dark')">
        🌙 Dark Mode
      </button>
    </div>

    <div class="bsp-body">
      <div class="bsp-stage" :class="`surface-${surface}`">
        <div class="bsp-surface-switch">
          <span>{{ isKo ? '문서/슬라이드 배경 테스트:' : 'Test Document Surface:' }}</span>
          <button type="button" :class="{ active: surface === 'light' }" @click="surface = 'light'">{{ isKo ? '일반 흰색 페이지' : 'White Page' }}</button>
          <button type="button" :class="{ active: surface === 'shaded' }" @click="surface = 'shaded'">{{ isKo ? '색상/음영 표 셀' : 'Shaded Table Cell' }}</button>
          <button type="button" :class="{ active: surface === 'dark' }" @click="surface = 'dark'">{{ isKo ? '다크 모드 슬라이드' : 'Dark Slide' }}</button>
        </div>

        <div class="bsp-mol-card" :class="{ 'has-white-box': activePreset === 'whitebg' }">
          <svg viewBox="0 0 265 200" class="bsp-svg">
            <g :stroke="bondStroke" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" fill="none">
              <polygon points="100,74 131,92 131,128 100,146 69,128 69,92" />
              <line x1="100" y1="81" x2="125" y2="96" />
              <line x1="125" y1="124" x2="100" y2="139" />
              <line x1="75" y1="124" x2="75" y2="96" />
              <line x1="131" y1="92" x2="162" y2="74" />
              <line x1="159" y1="75" x2="159" y2="43" />
              <line x1="165" y1="75" x2="165" y2="43" />
              <line x1="162" y1="74" x2="186" y2="88" />
              <line x1="131" y1="128" x2="153" y2="141" />
              <line x1="171" y1="141" x2="193" y2="128" />
              <line x1="190" y1="127" x2="190" y2="159" />
              <line x1="196" y1="127" x2="196" y2="159" />
              <line x1="193" y1="128" x2="224" y2="110" />
            </g>
            <g font-family="sans-serif" font-size="13.5" font-weight="700" :fill="heteroFill" text-anchor="middle">
              <text x="162" y="37">O</text>
              <text x="200" y="97">OH</text>
              <text x="162" y="151">O</text>
              <text x="193" y="174">O</text>
            </g>
          </svg>
        </div>
      </div>

      <div class="bsp-info">
        <h4>{{ presetInfo.title }}</h4>
        <p>{{ presetInfo.desc }}</p>
        <div class="bsp-equiv">
          <span>{{ isKo ? '동일한 엑셀 수식 표현:' : 'Equivalent Excel Formula:' }}</span>
          <code>{{ presetInfo.formula }}</code>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.bsp-tabs {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 8px;
  margin-bottom: 16px;
}
@media (max-width: 640px) {
  .bsp-tabs {
    grid-template-columns: repeat(2, 1fr);
  }
}
.bsp-tabs button {
  padding: 9px 10px;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font-weight: 600;
  font-size: 0.85rem;
  cursor: pointer;
  transition: all 0.15s;
}
.bsp-tabs button.active {
  background: var(--vp-c-brand-soft);
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}
.bsp-body {
  display: grid;
  grid-template-columns: 1.1fr 1fr;
  gap: 18px;
  align-items: center;
}
@media (max-width: 768px) {
  .bsp-body {
    grid-template-columns: 1fr;
  }
}
.bsp-stage {
  border-radius: 10px;
  padding: 14px;
  min-height: 210px;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  align-items: center;
  border: 1px solid var(--vp-c-divider);
  transition: background 0.25s;
}
.bsp-stage.surface-light {
  background: #ffffff;
  color: #334155;
}
.bsp-stage.surface-shaded {
  background: linear-gradient(135deg, #dbeafe 0%, #ede9fe 100%);
  color: #1e293b;
}
.bsp-stage.surface-dark {
  background: #0f172a;
  color: #e2e8f0;
}
.bsp-surface-switch {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 6px;
  font-size: 0.75rem;
  width: 100%;
  justify-content: center;
  margin-bottom: 10px;
}
.bsp-surface-switch button {
  font-size: 0.72rem;
  padding: 2px 8px;
  border-radius: 999px;
  border: 1px solid rgba(148, 163, 184, 0.4);
  background: rgba(255, 255, 255, 0.15);
  color: inherit;
  cursor: pointer;
}
.bsp-surface-switch button.active {
  font-weight: 700;
  border-color: #0d9488;
  background: rgba(20, 184, 166, 0.2);
}
.bsp-mol-card {
  padding: 10px 18px;
  border-radius: 8px;
  transition: background 0.2s;
}
.bsp-mol-card.has-white-box {
  background: #ffffff;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
}
.bsp-svg {
  width: 210px;
  height: auto;
}
.bsp-info h4 {
  margin: 0 0 8px 0;
  font-size: 1rem;
  color: var(--vp-c-text-1);
}
.bsp-info p {
  margin: 0 0 12px 0;
  font-size: 0.88rem;
  line-height: 1.55;
  color: var(--vp-c-text-2);
}
.bsp-equiv {
  font-size: 0.82rem;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
</style>
