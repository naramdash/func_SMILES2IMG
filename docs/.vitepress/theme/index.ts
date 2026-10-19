import { h, computed } from 'vue'
import DefaultTheme from 'vitepress/theme'
import { useData, withBase } from 'vitepress'
import type { Theme } from 'vitepress'
import FormulaBuilder from './components/FormulaBuilder.vue'
import BatchStylePreview from './components/BatchStylePreview.vue'
import OfficeAppTabs from './components/OfficeAppTabs.vue'
import QuickInstallBox from './components/QuickInstallBox.vue'
import './custom.css'

export default {
  extends: DefaultTheme,
  Layout() {
    const { lang } = useData()
    const isKo = computed(() => lang.value.startsWith('ko'))

    return h(DefaultTheme.Layout, null, {
      'home-hero-info-before': () =>
        h('div', { class: 's2i-hero-top-pill-wrap' }, [
          isKo.value
            ? h('div', { class: 's2i-lang-pill-row' }, [
                h(
                  'span',
                  { class: 's2i-hero-badge' },
                  '🧪 화학·제약 연구원용 무료 오피스 애드인 (func_SMILES2IMG)'
                ),
                h(
                  'a',
                  { href: withBase('/'), class: 's2i-lang-pill-sub' },
                  '🇺🇸 English Guide ➔'
                )
              ])
            : h(
                'a',
                { href: withBase('/ko/'), class: 's2i-lang-pill' },
                '🇰🇷 한국어 가이드 및 1분 설치 안내 바로가기 (Korean Guide) ➔'
              )
        ]),
      'home-features-before': () =>
        h('div', { class: 's2i-home-showcase-container' }, [h(OfficeAppTabs)])
    })
  },
  enhanceApp({ app }) {
    app.component('FormulaBuilder', FormulaBuilder)
    app.component('BatchStylePreview', BatchStylePreview)
    app.component('OfficeAppTabs', OfficeAppTabs)
    app.component('QuickInstallBox', QuickInstallBox)
  }
} satisfies Theme
