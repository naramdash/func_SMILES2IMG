import { h } from 'vue'
import DefaultTheme from 'vitepress/theme'
import type { Theme } from 'vitepress'
import { Icon } from '@iconify/vue'
import { registerIcons } from './icons'
import AppBadge from './components/AppBadge.vue'
import FormulaTable from './components/FormulaTable.vue'
import OfficeOverviewCta from './components/OfficeOverviewCta.vue'
import PlaybookTocCards from './components/PlaybookTocCards.vue'
import QuickInstallBox from './components/QuickInstallBox.vue'
import FeatureMedia from './components/FeatureMedia.vue'
import HeroDemoVideo from './components/HeroDemoVideo.vue'
import PlaybookPreview from './components/PlaybookPreview.vue'
import ImageLightbox from './components/ImageLightbox.vue'
import SidebarToggle from './components/SidebarToggle.vue'
import LangHintPopover from './components/LangHintPopover.vue'
import './custom.css'

registerIcons()

export default {
  extends: DefaultTheme,
  Layout() {
    return h(DefaultTheme.Layout, null, {
      'home-hero-image': () => h(HeroDemoVideo),
      'home-hero-after': () => h(QuickInstallBox),
      'nav-bar-content-before': () => h(SidebarToggle),
      'layout-bottom': () => [h(ImageLightbox), h(LangHintPopover)]
    })
  },
  enhanceApp({ app }) {
    app.component('Icon', Icon)
    app.component('AppBadge', AppBadge)
    app.component('FormulaTable', FormulaTable)
    app.component('OfficeOverviewCta', OfficeOverviewCta)
    app.component('PlaybookTocCards', PlaybookTocCards)
    app.component('QuickInstallBox', QuickInstallBox)
    app.component('FeatureMedia', FeatureMedia)
    app.component('PlaybookPreview', PlaybookPreview)
  }
} satisfies Theme
