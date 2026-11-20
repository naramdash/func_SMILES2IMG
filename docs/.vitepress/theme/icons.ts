import { addIcon } from '@iconify/vue'

export function registerIcons() {
  addIcon('lucide:panel-left-close', {
    body: '<g fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M9 3v18m7-6l-3-3l3-3"/></g>',
    width: 24,
    height: 24
  })
  addIcon('lucide:panel-left-open', {
    body: '<g fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M9 3v18m5-12l3 3l-3 3"/></g>',
    width: 24,
    height: 24
  })
  addIcon('lucide:copy', {
    body: '<g fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></g>',
    width: 24,
    height: 24
  })
  addIcon('lucide:check', {
    body: '<path fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20 6L9 17l-5-5"/>',
    width: 24,
    height: 24
  })
}
