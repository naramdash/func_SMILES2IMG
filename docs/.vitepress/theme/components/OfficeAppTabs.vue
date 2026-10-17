<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))
const activeApp = ref<'excel' | 'word' | 'ppt'>('excel')
</script>

<template>
  <div class="s2i-widget oat-root">
    <div class="oat-tabs">
      <button type="button" :class="{ active: activeApp === 'excel' }" @click="activeApp = 'excel'">
        📗 Microsoft Excel
      </button>
      <button type="button" :class="{ active: activeApp === 'word' }" @click="activeApp = 'word'">
        📘 Microsoft Word
      </button>
      <button type="button" :class="{ active: activeApp === 'ppt' }" @click="activeApp = 'ppt'">
        📙 Microsoft PowerPoint
      </button>
    </div>

    <div v-if="activeApp === 'excel'" class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/excel-showcase.png')" alt="SMILES for Excel" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>⚡ {{ isKo ? '수식으로 수백 개 구조식 자동 생성' : 'Auto-Render Hundreds of Structures' }}</strong>
          <span>{{ isKo ? '=SMILES2IMG(A2), =NAME2IMG("Aspirin") 입력 후 아래로 드래그' : 'Type =SMILES2IMG(A2) or =NAME2IMG("Aspirin") and drag down' }}</span>
        </div>
        <div class="oat-point">
          <strong>🔬 {{ isKo ? '분자량 · Lipinski 약물성 · SAR 유사도' : 'MW, Lipinski Ro5 & SAR Screening' }}</strong>
          <span>{{ isKo ? '=SMILES2MW, =SMILES2LIPINSKI, =SMILES.MATCH, =SMILES.SIMILARITY' : '0ms local =SMILES2MW, =SMILES2LIPINSKI, =SMILES.MATCH' }}</span>
        </div>
        <div class="oat-point">
          <strong>📂 {{ isKo ? '대용량 .sdf 파일 가져오기 / 내보내기' : 'Native .sdf Database Import & Export' }}</strong>
          <span>{{ isKo ? '.sdf 파일을 열면 구조식과 물성 수식이 걸린 엑셀 표 자동 완성' : 'Open .sdf files directly into formatted Excel tables with live formulas' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/excel' : '/guide/excel')" class="oat-btn primary">
          {{ isKo ? '📗 Excel 기초 가이드 보기 ➔' : '📗 Excel Basics Guide ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/excel-cheminformatics' : '/guide/excel-cheminformatics')" class="oat-btn">
          {{ isKo ? '🔬 물성 · SAR 스크리닝 · SDF 심화 ➔' : '🔬 SAR, Lipinski & SDF Guide ➔' }}
        </a>
      </div>
    </div>

    <div v-else-if="activeApp === 'word'" class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/word-showcase.png')" alt="SMILES for Word" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>📝 {{ isKo ? '문장 흐름에 맞춘 인라인 구조식' : 'Inline Structures Flowing with Text' }}</strong>
          <span>{{ isKo ? '글자처럼 줄바꿈되며 표(Table) 셀 안에서도 깔끔하게 정렬' : 'Behaves like text characters inside paragraphs and table cells' }}</span>
        </div>
        <div class="oat-point">
          <strong>✏️ {{ isKo ? '표 레이아웃 깨짐 없는 더블클릭 수정' : 'Double-Click Edit Without Breaking Tables' }}</strong>
          <span>{{ isKo ? '표 안의 그림을 더블클릭해 고쳐도 열 너비와 크기가 그대로 유지' : 'Updates structures in place while preserving table column widths' }}</span>
        </div>
        <div class="oat-point">
          <strong>🎨 {{ isKo ? '투고 학술지(ACS 흑백) 원클릭 일괄 변환' : '1-Click ACS Journal B&W Conversion' }}</strong>
          <span>{{ isKo ? '[Batch Style] 한 번으로 논문 내 모든 구조식을 출판 규격으로 통일' : 'Unify all manuscript structures to ACS monochrome in one click' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/word' : '/guide/word')" class="oat-btn primary">
          {{ isKo ? '📘 Word 논문·보고서 가이드 보기 ➔' : '📘 Word Manuscript Guide ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/publishing' : '/guide/publishing')" class="oat-btn">
          {{ isKo ? '🎨 출판 스타일 · SVG 벡터 · 물성 표 복사 ➔' : '🎨 Batch Style, SVG & Table Export ➔' }}
        </a>
      </div>
    </div>

    <div v-else class="oat-panel">
      <div class="oat-img-wrap">
        <img :src="withBase('/powerpoint-showcase.png')" alt="SMILES for PowerPoint" />
      </div>
      <div class="oat-points">
        <div class="oat-point">
          <strong>🌙 {{ isKo ? '다크 모드 슬라이드 투명 배경 최적화' : 'Dark Mode & Transparent Backgrounds' }}</strong>
          <span>{{ isKo ? '어두운 발표 템플릿 위에서도 흰 네모 테두리 없이 깔끔하게 표시' : 'Alpha-transparent PNGs & white ink bonds for dark slide decks' }}</span>
        </div>
        <div class="oat-point">
          <strong>🧬 {{ isKo ? '발표 중 마우스로 360° 돌려보는 3D 단백질' : 'Live 360°-Rotatable 3D Proteins' }}</strong>
          <span>{{ isKo ? 'PDB 코드로 단백질을 불러와 슬라이드 쇼 중 실시간 회전 시연' : 'Embed interactive .glb protein models directly onto your slides' }}</span>
        </div>
        <div class="oat-point">
          <strong>📋 {{ isKo ? '슬라이드 불릿 목록 & 포스터용 SVG 복사' : 'Slide Bullet Lists & Poster SVG Vectors' }}</strong>
          <span>{{ isKo ? '분자 물성을 글머리 기호로 붙여넣거나 학회 포스터용 무손실 벡터 복사' : 'Paste property bullet lists or copy razor-sharp SVG vectors for posters' }}</span>
        </div>
      </div>
      <div class="oat-links">
        <a :href="withBase(isKo ? '/ko/guide/powerpoint' : '/guide/powerpoint')" class="oat-btn primary">
          {{ isKo ? '📙 PowerPoint 발표 가이드 보기 ➔' : '📙 PowerPoint Slide Guide ➔' }}
        </a>
        <a :href="withBase(isKo ? '/ko/guide/molstar' : '/guide/molstar')" class="oat-btn">
          {{ isKo ? '🧬 Mol* 3D 단백질 뷰어 가이드 ➔' : '🧬 Mol* 3D Protein Viewer Guide ➔' }}
        </a>
      </div>
    </div>
  </div>
</template>

<style scoped>
.oat-tabs {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 8px;
  margin-bottom: 16px;
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
