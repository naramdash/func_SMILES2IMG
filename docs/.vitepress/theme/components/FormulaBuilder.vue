<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData } from 'vitepress'

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const inputType = ref<'smiles' | 'name' | 'cas' | 'ket' | 'inchi'>('smiles')
const isFloat = ref(false)
const bgMode = ref<'' | 'white' | '#1e1e1e' | '#fff9db'>('')
const colorMode = ref<'colored' | 'bw' | 'white'>('colored')
const highlightMode = ref<'' | 'hl:c1ccccc1' | 'hl:C(=O)O'>('')

const optNum = ref(false)
const optStereo = ref(false)
const optH = ref(false)
const optAllC = ref(false)
const optFlip = ref(false)
const optFlipY = ref(false)
const optRot90 = ref(false)

const copied = ref(false)

function selectDarkPreset() {
  bgMode.value = '#1e1e1e'
  colorMode.value = 'white'
}

function selectAcsPreset() {
  bgMode.value = ''
  colorMode.value = 'bw'
  highlightMode.value = ''
}

function selectHighlightPreset() {
  bgMode.value = ''
  colorMode.value = 'colored'
  highlightMode.value = 'hl:c1ccccc1'
}

function selectNmrPreset() {
  bgMode.value = ''
  colorMode.value = 'colored'
  optNum.value = true
  optStereo.value = true
}

function resetAll() {
  inputType.value = 'smiles'
  isFloat.value = false
  bgMode.value = ''
  colorMode.value = 'colored'
  highlightMode.value = ''
  optNum.value = false
  optStereo.value = false
  optH.value = false
  optAllC.value = false
  optFlip.value = false
  optFlipY.value = false
  optRot90.value = false
}

const funcName = computed(() => {
  const base =
    inputType.value === 'name' || inputType.value === 'cas'
      ? 'NAME2IMG'
      : inputType.value === 'ket'
      ? 'KET2IMG'
      : inputType.value === 'inchi'
      ? 'INCHI2IMG'
      : 'SMILES2IMG'
  return isFloat.value ? `${base}.FLOAT` : base
})

const firstArg = computed(() => {
  if (inputType.value === 'name') return '"Aspirin"'
  if (inputType.value === 'cas') return '"50-78-2"'
  return 'A2'
})

const styleTokens = computed(() => {
  const tokens: string[] = []
  if (colorMode.value === 'bw') tokens.push('bw')
  if (colorMode.value === 'white') tokens.push('white')
  if (highlightMode.value) tokens.push(highlightMode.value)
  if (optNum.value) tokens.push('num')
  if (optStereo.value) tokens.push('stereo')
  if (optH.value) tokens.push('h')
  if (optAllC.value) tokens.push('all-c')
  if (optFlip.value) tokens.push('flip')
  if (optFlipY.value) tokens.push('flipy')
  if (optRot90.value) tokens.push('rot90')
  return tokens
})

const generatedFormula = computed(() => {
  const styleStr = styleTokens.value.join(' ')
  if (!bgMode.value && !styleStr) {
    return `=${funcName.value}(${firstArg.value})`
  }
  if (bgMode.value && !styleStr) {
    return `=${funcName.value}(${firstArg.value}, "${bgMode.value}")`
  }
  return `=${funcName.value}(${firstArg.value}, "${bgMode.value}", "${styleStr}")`
})

const bondColor = computed(() => {
  if (colorMode.value === 'white') return '#f8fafc'
  return '#1e293b'
})

const oxygenColor = computed(() => {
  if (colorMode.value === 'bw') return '#1e293b'
  if (colorMode.value === 'white') return '#f87171'
  return '#dc2626'
})

const previewBgStyle = computed(() => {
  if (bgMode.value === '#1e1e1e') return { background: '#1e1e1e', borderColor: '#334155' }
  if (bgMode.value === '#fff9db') return { background: '#fff9db', borderColor: '#fde047' }
  if (bgMode.value === 'white') return { background: '#ffffff', borderColor: '#cbd5e1' }
  return {
    background: 'repeating-conic-gradient(#f1f5f9 0% 25%, #ffffff 0% 50%) 50% / 16px 16px',
    borderColor: '#cbd5e1'
  }
})

const svgTransform = computed(() => {
  const parts: string[] = []
  if (optRot90.value) parts.push('rotate(90deg)')
  if (optFlip.value) parts.push('scaleX(-1)')
  if (optFlipY.value) parts.push('scaleY(-1)')
  return parts.join(' ')
})

async function copyFormula() {
  try {
    await navigator.clipboard.writeText(generatedFormula.value)
    copied.value = true
    setTimeout(() => {
      copied.value = false
    }, 1600)
  } catch {
    // ignore clipboard errors
  }
}
</script>

<template>
  <div class="s2i-widget fb-root">
    <div class="s2i-widget-header">
      <div class="s2i-widget-title">
        <span>🧪 {{ isKo ? '인터랙티브 엑셀 수식 생성기 (클릭해서 조합해 보세요!)' : 'Interactive Excel Formula Builder (Click to Customize!)' }}</span>
      </div>
      <div class="fb-presets">
        <span class="fb-preset-label">{{ isKo ? '빠른 프리셋:' : 'Quick Presets:' }}</span>
        <button type="button" class="fb-chip" @click="resetAll">{{ isKo ? '초기화' : 'Default' }}</button>
        <button type="button" class="fb-chip" @click="selectAcsPreset">{{ isKo ? '📄 ACS 논문 흑백' : '📄 ACS B&W' }}</button>
        <button type="button" class="fb-chip" @click="selectHighlightPreset">{{ isKo ? '🔍 벤젠 하이라이트' : '🔍 Highlight Ring' }}</button>
        <button type="button" class="fb-chip" @click="selectNmrPreset">{{ isKo ? '🔢 NMR 원자번호' : '🔢 NMR Numbers' }}</button>
        <button type="button" class="fb-chip" @click="selectDarkPreset">{{ isKo ? '🌙 다크 모드' : '🌙 Dark Mode' }}</button>
      </div>
    </div>

    <!-- Formula Output Bar -->
    <div class="fb-formula-bar">
      <div class="fb-formula-code">
        <code>{{ generatedFormula }}</code>
      </div>
      <button type="button" class="fb-copy-btn" @click="copyFormula">
        {{ copied ? (isKo ? '✓ 복사됨!' : '✓ Copied!') : (isKo ? '📋 수식 복사' : '📋 Copy Formula') }}
      </button>
    </div>

    <div class="fb-layout">
      <!-- Left Controls -->
      <div class="fb-controls">
        <div class="fb-group">
          <label class="fb-label">{{ isKo ? '1. 내 입력 데이터 종류' : '1. Input Data Type' }}</label>
          <div class="fb-btn-group">
            <button type="button" :class="{ active: inputType === 'smiles' }" @click="inputType = 'smiles'">SMILES (A2)</button>
            <button type="button" :class="{ active: inputType === 'name' }" @click="inputType = 'name'">{{ isKo ? '물질명 ("Aspirin")' : 'Name ("Aspirin")' }}</button>
            <button type="button" :class="{ active: inputType === 'cas' }" @click="inputType = 'cas'">CAS ("50-78-2")</button>
            <button type="button" :class="{ active: inputType === 'ket' }" @click="inputType = 'ket'">{{ isKo ? '반응식 KET (A2)' : 'Reaction KET' }}</button>
            <button type="button" :class="{ active: inputType === 'inchi' }" @click="inputType = 'inchi'">InChI (A2)</button>
          </div>
        </div>

        <div class="fb-group">
          <label class="fb-label">{{ isKo ? '2. 내 엑셀 버전' : '2. Excel Version' }}</label>
          <div class="fb-btn-group">
            <button type="button" :class="{ active: !isFloat }" @click="isFloat = false">
              {{ isKo ? 'Microsoft 365 / Excel 2024+ (인셀 이미지)' : 'Microsoft 365 / Excel 2024+ (In-Cell)' }}
            </button>
            <button type="button" :class="{ active: isFloat }" @click="isFloat = true">
              {{ isKo ? 'Excel 2016 / 2019 / 2021 (.FLOAT 도형)' : 'Excel 2016–2021 (.FLOAT Shape)' }}
            </button>
          </div>
        </div>

        <div class="fb-group">
          <label class="fb-label">{{ isKo ? '3. 배경색 및 선 색상 모드' : '3. Background & Color Mode' }}</label>
          <div class="fb-btn-group">
            <button type="button" :class="{ active: bgMode === '' }" @click="bgMode = ''">{{ isKo ? '투명 배경 (기본)' : 'Transparent BG' }}</button>
            <button type="button" :class="{ active: bgMode === 'white' }" @click="bgMode = 'white'">{{ isKo ? '흰색 배경 ("white")' : 'White BG ("white")' }}</button>
            <button type="button" :class="{ active: bgMode === '#1e1e1e' }" @click="bgMode = '#1e1e1e'">{{ isKo ? '어두운 배경 ("#1e1e1e")' : 'Dark BG ("#1e1e1e")' }}</button>
            <button type="button" :class="{ active: bgMode === '#fff9db' }" @click="bgMode = '#fff9db'">{{ isKo ? '노란색 ("#fff9db")' : 'Yellow ("#fff9db")' }}</button>
          </div>
          <div class="fb-btn-group" style="margin-top: 6px;">
            <button type="button" :class="{ active: colorMode === 'colored' }" @click="colorMode = 'colored'">{{ isKo ? '🎨 표준 원소 컬러' : '🎨 Standard Color' }}</button>
            <button type="button" :class="{ active: colorMode === 'bw' }" @click="colorMode = 'bw'">{{ isKo ? '✒️ 논문용 흑백 (bw)' : '✒️ Monochrome (bw)' }}</button>
            <button type="button" :class="{ active: colorMode === 'white' }" @click="colorMode = 'white'">{{ isKo ? '🌙 흰색 결합선 (white)' : '🌙 White Ink (white)' }}</button>
          </div>
        </div>

        <div class="fb-group">
          <label class="fb-label">{{ isKo ? '4. 부분구조 하이라이트 (hl:)' : '4. Substructure Highlight (hl:)' }}</label>
          <div class="fb-btn-group">
            <button type="button" :class="{ active: highlightMode === '' }" @click="highlightMode = ''">{{ isKo ? '없음' : 'None' }}</button>
            <button type="button" :class="{ active: highlightMode === 'hl:c1ccccc1' }" @click="highlightMode = 'hl:c1ccccc1'">
              {{ isKo ? '벤젠 고리 강조 (hl:c1ccccc1)' : 'Benzene Ring (hl:c1ccccc1)' }}
            </button>
            <button type="button" :class="{ active: highlightMode === 'hl:C(=O)O' }" @click="highlightMode = 'hl:C(=O)O'">
              {{ isKo ? '카복실기 강조 (hl:C(=O)O)' : 'Carboxyl Group (hl:C(=O)O)' }}
            </button>
          </div>
        </div>

        <div class="fb-group">
          <label class="fb-label">{{ isKo ? '5. 추가 라벨 & 회전/반전 (중복 선택 가능)' : '5. Labels & Geometry (Multi-select)' }}</label>
          <div class="fb-check-grid">
            <label class="fb-check"><input type="checkbox" v-model="optNum" /> <span><code>num</code> ({{ isKo ? '원자 번호' : 'Atom #' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optStereo" /> <span><code>stereo</code> ({{ isKo ? '입체 R/S' : 'Stereo R/S' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optH" /> <span><code>h</code> ({{ isKo ? '수소 펼침' : 'Unfold H' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optAllC" /> <span><code>all-c</code> ({{ isKo ? '탄소 기호 C' : 'Show all C' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optFlip" /> <span><code>flip</code> ({{ isKo ? '좌우 반전' : 'Flip H' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optFlipY" /> <span><code>flipy</code> ({{ isKo ? '상하 반전' : 'Flip V' }})</span></label>
            <label class="fb-check"><input type="checkbox" v-model="optRot90" /> <span><code>rot90</code> ({{ isKo ? '90° 회전' : 'Rotate 90°' }})</span></label>
          </div>
        </div>
      </div>

      <!-- Right Live Preview -->
      <div class="fb-preview-col">
        <div class="fb-preview-label">{{ isKo ? '실시간 셀 렌더링 미리보기' : 'Live Cell Rendering Preview' }}</div>
        <div class="fb-canvas" :style="previewBgStyle">
          <svg viewBox="0 0 260 210" class="fb-svg" :style="{ transform: svgTransform }">
            <!-- Highlight: Benzene ring -->
            <g v-if="highlightMode === 'hl:c1ccccc1'" stroke="#f59e0b" stroke-width="12" stroke-linecap="round" stroke-linejoin="round" fill="rgba(251, 191, 36, 0.22)">
              <polygon points="80,85 115,65 150,85 150,125 115,145 80,125" />
            </g>
            <!-- Highlight: Carboxyl group -->
            <g v-if="highlightMode === 'hl:C(=O)O'" stroke="#ec4899" stroke-width="12" stroke-linecap="round" stroke-linejoin="round" fill="none">
              <line x1="150" y1="85" x2="185" y2="65" />
              <line x1="185" y1="65" x2="185" y2="30" />
              <line x1="185" y1="65" x2="218" y2="85" />
              <circle cx="185" cy="25" r="10" fill="rgba(236, 72, 153, 0.25)" stroke="none" />
              <circle cx="224" cy="88" r="12" fill="rgba(236, 72, 153, 0.25)" stroke="none" />
            </g>

            <!-- Benzene Ring Bonds -->
            <g :stroke="bondColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" fill="none">
              <polygon points="80,85 115,65 150,85 150,125 115,145 80,125" />
              <!-- Aromatic inner double bonds -->
              <line x1="87" y1="89" x2="115" y2="73" />
              <line x1="143" y1="89" x2="143" y2="121" />
              <line x1="115" y1="137" x2="87" y2="121" />

              <!-- Carboxyl attachment at C2 (150,85) -> C7 (185,65) -->
              <line x1="150" y1="85" x2="185" y2="65" />
              <!-- C7=O double bond -->
              <line x1="182" y1="65" x2="182" y2="36" />
              <line x1="188" y1="65" x2="188" y2="36" />
              <!-- C7-OH single bond -->
              <line x1="185" y1="65" x2="212" y2="81" />

              <!-- Ester attachment at C3 (150,125) -> O (185,145) -> C(=O)CH3 -->
              <line x1="150" y1="125" x2="177" y2="140" />
              <line x1="193" y1="140" x2="220" y2="125" />
              <line x1="217" y1="125" x2="217" y2="96" />
              <line x1="223" y1="125" x2="223" y2="96" />
              <line x1="220" y1="125" x2="248" y2="142" />
            </g>

            <!-- Explicit Hydrogens when optH is enabled -->
            <g v-if="optH" :stroke="bondColor" stroke-width="1.8" :fill="bondColor" font-size="11" font-family="sans-serif" text-anchor="middle">
              <line x1="115" y1="65" x2="115" y2="43" />
              <text x="115" y="39" stroke="none">H</text>
              <line x1="80" y1="85" x2="58" y2="72" />
              <text x="50" y="71" stroke="none">H</text>
              <line x1="80" y1="125" x2="58" y2="138" />
              <text x="50" y="144" stroke="none">H</text>
              <line x1="115" y1="145" x2="115" y2="168" />
              <text x="115" y="180" stroke="none">H</text>
            </g>

            <!-- Explicit Carbons when optAllC is enabled -->
            <g v-if="optAllC" font-size="11" font-weight="700" font-family="sans-serif" text-anchor="middle" :fill="bondColor">
              <circle v-for="[cx, cy] in [[80,85],[115,65],[150,85],[150,125],[115,145],[80,125],[185,65],[220,125],[248,142]]" :key="`${cx}-${cy}`" :cx="cx" :cy="cy" r="7" :fill="bgMode === '#1e1e1e' ? '#1e1e1e' : '#ffffff'" />
              <text x="80" y="89">C</text>
              <text x="115" y="69">C</text>
              <text x="150" y="89">C</text>
              <text x="150" y="129">C</text>
              <text x="115" y="149">C</text>
              <text x="80" y="129">C</text>
              <text x="185" y="69">C</text>
              <text x="220" y="129">C</text>
              <text x="248" y="146">{{ optH ? 'CH₃' : 'C' }}</text>
            </g>

            <!-- Heteroatoms (O, OH) -->
            <g font-family="sans-serif" font-size="14" font-weight="700" :fill="oxygenColor" text-anchor="middle">
              <text x="185" y="31">O</text>
              <text x="225" y="91">OH</text>
              <text x="185" y="150">O</text>
              <text x="220" y="91">O</text>
            </g>

            <!-- Stereo annotation when optStereo is enabled -->
            <g v-if="optStereo" font-family="sans-serif" font-size="11" font-style="italic" font-weight="700" fill="#0d9488">
              <text x="142" y="76">(R/S)</text>
            </g>

            <!-- Atom Numbers when optNum is enabled -->
            <g v-if="optNum" font-family="sans-serif" font-size="10" font-weight="700" fill="#2563eb">
              <text x="68" y="82">1</text>
              <text x="112" y="57">2</text>
              <text x="154" y="78">3</text>
              <text x="155" y="122">4</text>
              <text x="112" y="159">5</text>
              <text x="68" y="135">6</text>
              <text x="191" y="62">7</text>
              <text x="195" y="26">8</text>
              <text x="238" y="82">9</text>
              <text x="182" y="162">10</text>
              <text x="228" y="132">11</text>
              <text x="250" y="156">12</text>
            </g>
          </svg>
        </div>
        <div class="fb-caption">
          {{ isKo ? '💡 투명 배경(체크무늬)은 엑셀 셀에 들어갈 때 셀 배경색과 자연스럽게 합쳐집니다.' : '💡 Checkerboard indicates transparent background that blends with Excel cells.' }}
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.fb-presets {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 6px;
}
.fb-preset-label {
  font-size: 0.8rem;
  color: var(--vp-c-text-2);
  margin-right: 4px;
}
.fb-chip {
  font-size: 0.78rem;
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  cursor: pointer;
  transition: all 0.15s;
}
.fb-chip:hover {
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}
.fb-formula-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  background: var(--vp-c-bg);
  border: 2px solid var(--vp-c-brand-1);
  border-radius: 10px;
  padding: 10px 14px;
  margin-bottom: 18px;
}
.fb-formula-code {
  overflow-x: auto;
  font-size: 1rem;
  font-weight: 700;
}
.fb-formula-code code {
  background: transparent !important;
  color: var(--vp-c-brand-1);
  padding: 0;
}
.fb-copy-btn {
  flex-shrink: 0;
  background: var(--vp-c-brand-1);
  color: #ffffff;
  font-weight: 600;
  font-size: 0.85rem;
  padding: 6px 14px;
  border-radius: 8px;
  border: none;
  cursor: pointer;
  transition: opacity 0.15s;
}
.dark .fb-copy-btn {
  color: #0f172a;
}
.fb-copy-btn:hover {
  opacity: 0.9;
}
.fb-layout {
  display: grid;
  grid-template-columns: 1.35fr 0.85fr;
  gap: 20px;
  align-items: start;
}
@media (max-width: 768px) {
  .fb-layout {
    grid-template-columns: 1fr;
  }
}
.fb-group {
  margin-bottom: 14px;
}
.fb-label {
  display: block;
  font-size: 0.82rem;
  font-weight: 700;
  color: var(--vp-c-text-2);
  margin-bottom: 6px;
}
.fb-btn-group {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}
.fb-btn-group button {
  font-size: 0.8rem;
  padding: 5px 10px;
  border-radius: 6px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  cursor: pointer;
  transition: all 0.15s;
}
.fb-btn-group button.active {
  background: var(--vp-c-brand-soft);
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
  font-weight: 600;
}
.fb-check-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(145px, 1fr));
  gap: 8px;
}
.fb-check {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 0.82rem;
  cursor: pointer;
  padding: 4px 8px;
  border-radius: 6px;
  background: var(--vp-c-bg);
  border: 1px solid var(--vp-c-divider);
}
.fb-preview-col {
  display: flex;
  flex-direction: column;
  align-items: center;
}
.fb-preview-label {
  font-size: 0.82rem;
  font-weight: 700;
  color: var(--vp-c-text-2);
  margin-bottom: 8px;
}
.fb-canvas {
  width: 100%;
  min-height: 220px;
  border: 2px solid #cbd5e1;
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 12px;
  overflow: hidden;
  transition: background 0.2s, border-color 0.2s;
}
.fb-svg {
  width: 100%;
  max-width: 250px;
  height: auto;
  transition: transform 0.25s ease;
}
.fb-caption {
  font-size: 0.76rem;
  color: var(--vp-c-text-2);
  margin-top: 8px;
  text-align: center;
  line-height: 1.4;
}
</style>
