<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const activeApp = ref<'excel' | 'word' | 'ppt'>('excel')
</script>

<template>
  <div class="s2i-widget oat-root">
    <div class="oat-header">
      <h2 class="oat-heading">
        {{ isKo ? '🖥️ 오피스 프로그램별 실제 작동 화면' : '🖥️ See It in Action Across Microsoft Office' }}
      </h2>
      <span class="oat-hint">
        {{ isKo ? '탭을 클릭해 Excel · Word · PowerPoint 화면을 확인해 보세요' : 'Click a tab to preview Excel, Word, and PowerPoint workflows' }}
      </span>
    </div>

    <div class="oat-tabs" role="tablist" :aria-label="isKo ? 'Office 앱 선택' : 'Select Office App'">
      <button type="button" role="tab" :aria-selected="activeApp === 'excel'" :class="{ active: activeApp === 'excel' }" @click="activeApp = 'excel'">
        📗 Microsoft Excel
      </button>
      <button type="button" role="tab" :aria-selected="activeApp === 'word'" :class="{ active: activeApp === 'word' }" @click="activeApp = 'word'">
        📘 Microsoft Word
      </button>
      <button type="button" role="tab" :aria-selected="activeApp === 'ppt'" :class="{ active: activeApp === 'ppt' }" @click="activeApp = 'ppt'">
        📙 Microsoft PowerPoint
      </button>
    </div>

    <div v-if="activeApp === 'excel'" class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/excel-showcase.png')" :alt="isKo ? 'Microsoft Excel에서 실행 중인 func_SMILES2IMG 구조식 및 물성 표' : 'func_SMILES2IMG running in Microsoft Excel'" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>⚡ {{ isKo ? '화합물 이름·SMILES로 구조식 표 자동 완성' : 'Auto-Render Compound Tables in Seconds' }}</strong>
          <span>{{ isKo ? '시약 이름이나 SMILES를 적고 아래로 드래그하면 수백 행의 구조식이 즉시 그려집니다.' : 'Type compound names, CAS numbers, or SMILES and drag down across hundreds of rows.' }}</span>
        </div>
        <div class="oat-point">
          <strong>🔬 {{ isKo ? '외부 툴 없이 물성 계산 & SAR 스크리닝' : 'Instant Local Property & SAR Screening' }}</strong>
          <span>{{ isKo ? '분자량(MW), LogP, Lipinski 약물성 통과 여부와 특정 골격 포함 여부를 시트에서 바로 판별합니다.' : 'Compute MW, LogP, Lipinski Rule of 5 pass/fail, and substructure matches with 0ms lag.' }}</span>
        </div>
        <div class="oat-point">
          <strong>📂 {{ isKo ? '대용량 .sdf 화학 DB 불러오기 / 내보내기' : '1-Click .sdf Database Import & Export' }}</strong>
          <span>{{ isKo ? '.sdf 파일을 열면 구조식과 물성이 정리된 엑셀 시트가 자동으로 만들어집니다.' : 'Open .sdf files directly into formatted Excel tables or export sheets back to .sdf.' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/excel' : '/guide/excel')" class="oat-btn primary">
          {{ isKo ? '📗 Excel 레시피 1~5 (구조식 · 스타일 · 반응식) ➔' : '📗 Excel Recipes 1–5 (Tables & Styles) ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/excel-cheminformatics' : '/guide/excel-cheminformatics')" class="oat-btn">
          {{ isKo ? '🔬 Excel 레시피 6~10 (물성 · SAR · SDF) ➔' : '🔬 Excel Recipes 6–10 (SAR, Lipinski & SDF) ➔' }}
        </a>
      </div>
    </div>

    <div v-else-if="activeApp === 'word'" class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/word-showcase.png')" :alt="isKo ? 'Microsoft Word에서 실행 중인 func_SMILES2IMG 논문 원고' : 'func_SMILES2IMG running in Microsoft Word'" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>📝 {{ isKo ? '1편의 논문 원고를 6장에 걸쳐 완성' : '6-Chapter Continuous Manuscript Tutorial' }}</strong>
          <span>{{ isKo ? '서론 인라인 구조식부터 화합물 요약 표, 합성 Scheme, 투고 마무리까지 순서대로 따라 합니다.' : 'Build one real research manuscript from Ch. 1 prose figures through Ch. 6 ACS B&W submission.' }}</span>
        </div>
        <div class="oat-point">
          <strong>✏️ {{ isKo ? '표 레이아웃 깨짐 없는 더블클릭 제자리 수정' : 'Double-Click Edit Without Breaking Tables' }}</strong>
          <span>{{ isKo ? '표 안의 구조식을 더블클릭해 수정해도 열 너비와 문서 배치가 절대 틀어지지 않습니다.' : 'Re-edit any molecule in place while preserving exact table column widths and margins.' }}</span>
        </div>
        <div class="oat-point">
          <strong>🎨 {{ isKo ? '학술지 투고 규격(ACS 흑백) 원클릭 통일' : '1-Click ACS Journal B&W Conversion' }}</strong>
          <span>{{ isKo ? '클릭 한 번으로 논문 전체의 구조식을 학술지 흑백 선화로 일괄 변환합니다.' : 'Standardize every structure across your entire manuscript into ACS monochrome in one click.' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/word' : '/guide/word')" class="oat-btn primary">
          {{ isKo ? '📘 Word 6장 논문 튜토리얼 보기 ➔' : '📘 Word 6-Chapter Tutorial ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/publishing' : '/guide/publishing')" class="oat-btn">
          {{ isKo ? '🎨 출판 스타일 · SVG 벡터 · 물성 표 복사 ➔' : '🎨 Batch Style, SVG & Table Export ➔' }}
        </a>
      </div>
    </div>

    <div v-else class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/powerpoint-showcase.png')" :alt="isKo ? 'Microsoft PowerPoint 다크 테마 슬라이드에서 실행 중인 func_SMILES2IMG' : 'func_SMILES2IMG running in Microsoft PowerPoint dark slide'" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>🌙 {{ isKo ? '다크 모드 발표 슬라이드 투명 배경 최적화' : 'Dark Mode & Transparent Slide Backgrounds' }}</strong>
          <span>{{ isKo ? '어두운 발표 템플릿 위에서도 하얀 네모 여백 없이 깔끔한 흰색 선화로 표시됩니다.' : 'Alpha-transparent PNGs and crisp white-ink bonds tailored for dark presentation decks.' }}</span>
        </div>
        <div class="oat-point">
          <strong>🧬 {{ isKo ? '슬라이드 쇼 중 마우스로 돌려보는 3D 단백질' : 'Live 360°-Rotatable 3D Proteins on Slides' }}</strong>
          <span>{{ isKo ? 'PDB 코드로 단백질·리간드 결합 모델을 넣어 발표 중 실시간으로 360° 회전 시연합니다.' : 'Embed interactive 3D protein models directly onto slides and rotate live during talks.' }}</span>
        </div>
        <div class="oat-point">
          <strong>📋 {{ isKo ? '완성 슬라이드 갤러리 보고 3단계 따라 만들기' : 'Visual Slide Gallery ➔ 3-Step Follow-Along' }}</strong>
          <span>{{ isKo ? '원하는 슬라이드 디자인의 실제 작동 영상을 먼저 보고 3단계로 바로 제작합니다.' : 'Preview live-recorded slide decks first, then follow 3 steps to recreate each layout.' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/powerpoint' : '/guide/powerpoint')" class="oat-btn primary">
          {{ isKo ? '📙 PowerPoint 슬라이드 갤러리 보기 ➔' : '📙 PowerPoint Slide Gallery ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/molstar' : '/guide/molstar')" class="oat-btn">
          {{ isKo ? '🧬 Mol* 3D 단백질 뷰어 가이드 ➔' : '🧬 Mol* 3D Protein Viewer Guide ➔' }}
        </a>
      </div>
    </div>
  </div>
</template>

<style scoped>
.oat-header {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 14px;
}
.oat-heading {
  margin: 0 !important;
  padding: 0 !important;
  border: none !important;
  font-size: 1.25rem;
  font-weight: 800;
  color: var(--vp-c-text-1);
}
.oat-hint {
  font-size: 0.84rem;
  color: var(--vp-c-text-2);
}
.oat-tabs {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 8px;
  margin-bottom: 16px;
}
@media (max-width: 640px) {
  .oat-tabs {
    grid-template-columns: 1fr;
  }
}
.oat-tabs button {
  padding: 10px 12px;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font-weight: 700;
  font-size: 0.92rem;
  cursor: pointer;
  transition: all 0.15s;
}
.oat-tabs button:hover {
  border-color: var(--vp-c-brand-1);
}
.oat-tabs button.active {
  background: var(--vp-c-brand-soft);
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}
.oat-img-wrap img {
  width: 100%;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  margin-bottom: 14px;
}
.oat-points {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 12px;
  margin-bottom: 16px;
}
.oat-point {
  background: var(--vp-c-bg);
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.oat-point strong {
  font-size: 0.9rem;
  color: var(--vp-c-text-1);
}
.oat-point span {
  font-size: 0.82rem;
  color: var(--vp-c-text-2);
  line-height: 1.45;
}
.oat-links {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}
.oat-btn {
  display: inline-block;
  padding: 8px 16px;
  border-radius: 8px;
  font-size: 0.88rem;
  font-weight: 600;
  text-decoration: none !important;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1) !important;
  transition: all 0.15s;
}
.oat-btn.primary {
  background: var(--vp-c-brand-1);
  border-color: var(--vp-c-brand-1);
  color: #ffffff !important;
}
.dark .oat-btn.primary {
  color: #0f172a !important;
}
.oat-btn:hover {
  transform: translateY(-1px);
}
</style>
