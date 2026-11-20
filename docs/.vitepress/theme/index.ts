import { h } from 'vue'
import DefaultTheme from 'vitepress/theme'
import type { Theme } from 'vitepress'
import { Icon } from '@iconify/vue'
import { registerIcons } from './icons'
import FormulaBuilder from './components/FormulaBuilder.vue'
import QuickInstallBox from './components/QuickInstallBox.vue'
import FeatureMedia from './components/FeatureMedia.vue'
import HeroDemoVideo from './components/HeroDemoVideo.vue'
import PlaybookPreview from './components/PlaybookPreview.vue'
import SidebarToggle from './components/SidebarToggle.vue'
import './custom.css'

registerIcons()

export default {
  extends: DefaultTheme,
  Layout() {
    return h(DefaultTheme.Layout, null, {
      'home-hero-image': () => h(HeroDemoVideo),
      'home-hero-after': () => h(QuickInstallBox),
      'nav-bar-content-before': () => h(SidebarToggle)
    })
  },
  enhanceApp({ app }) {
    app.component('Icon', Icon)
    app.component('FormulaBuilder', FormulaBuilder)
    app.component('FeatureMedia', FeatureMedia)
    app.component('PlaybookPreview', PlaybookPreview)
  }
} satisfies Theme
