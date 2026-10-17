import DefaultTheme from 'vitepress/theme'
import type { Theme } from 'vitepress'
import FormulaBuilder from './components/FormulaBuilder.vue'
import BatchStylePreview from './components/BatchStylePreview.vue'
import OfficeAppTabs from './components/OfficeAppTabs.vue'
import './custom.css'

export default {
  extends: DefaultTheme,
  enhanceApp({ app }) {
    app.component('FormulaBuilder', FormulaBuilder)
    app.component('BatchStylePreview', BatchStylePreview)
    app.component('OfficeAppTabs', OfficeAppTabs)
  }
} satisfies Theme
