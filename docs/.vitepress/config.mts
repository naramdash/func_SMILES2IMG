import { defineConfig } from 'vitepress'

export default defineConfig({
  title: 'func_SMILES2IMG',
  description: 'Free & Open-Source Chemical Structure Add-in Suite for Microsoft Excel, Word, and PowerPoint',
  base: '/func_SMILES2IMG/',
  cleanUrls: true,
  lastUpdated: true,

  head: [],

  locales: {
    root: {
      label: 'English',
      lang: 'en',
      title: 'func_SMILES2IMG',
      description: 'Free & Open-Source Chemical Structure Add-in Suite for Microsoft Excel, Word, and PowerPoint',
      themeConfig: {
        nav: [
          { text: 'Guide', link: '/guide/getting-started', activeMatch: '^/guide/' },
          { text: 'Reference', link: '/reference/formula', activeMatch: '^/reference/' },
          { text: 'v1.7.0-preview.5', link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.0-preview.5' }
        ],
        sidebar: [
          {
            text: 'Getting Started',
            items: [
              { text: 'Overview & Installation', link: '/guide/getting-started' }
            ]
          },
          {
            text: 'Editors & 3D Viewers',
            items: [
              { text: 'Ketcher 2D Editor', link: '/guide/ketcher' },
              { text: 'Mol* 3D Viewer', link: '/guide/molstar' }
            ]
          },
          {
            text: 'Office Applications',
            items: [
              { text: 'Microsoft Excel', link: '/guide/excel' },
              { text: 'Microsoft Word', link: '/guide/word' },
              { text: 'Microsoft PowerPoint', link: '/guide/powerpoint' }
            ]
          },
          {
            text: 'Reference',
            items: [
              { text: 'Excel Formulas & Options', link: '/reference/formula' }
            ]
          },
          {
            text: 'Help & Support',
            items: [
              { text: 'Troubleshooting & FAQ', link: '/guide/faq' }
            ]
          }
        ],
        editLink: {
          pattern: 'https://github.com/naramdash/func_SMILES2IMG/edit/dev/docs/:path',
          text: 'Edit this page on GitHub'
        },
        footer: {
          message: 'Released under the MIT License.',
          copyright: 'Copyright © 2026 KimJuho. Powered by EPAM Indigo, Ketcher & Mol*.'
        }
      }
    },
    ko: {
      label: '한국어',
      lang: 'ko',
      link: '/ko/',
      title: 'func_SMILES2IMG',
      description: 'Microsoft Excel, Word, PowerPoint용 무료 오픈소스 화학 구조식 렌더러 & 에디터 추가 기능',
      themeConfig: {
        nav: [
          { text: '가이드', link: '/ko/guide/getting-started', activeMatch: '^/ko/guide/' },
          { text: '수식 레퍼런스', link: '/ko/reference/formula', activeMatch: '^/ko/reference/' },
          { text: 'v1.7.0-preview.5', link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.0-preview.5' }
        ],
        sidebar: [
          {
            text: '시작하기',
            items: [
              { text: '개요 및 빠른 설치', link: '/ko/guide/getting-started' }
            ]
          },
          {
            text: '편집기 & 3D 뷰어',
            items: [
              { text: 'Ketcher 2D 에디터', link: '/ko/guide/ketcher' },
              { text: 'Mol* 3D 고분자 뷰어', link: '/ko/guide/molstar' }
            ]
          },
          {
            text: '오피스 프로그램별 가이드',
            items: [
              { text: 'Microsoft Excel 가이드', link: '/ko/guide/excel' },
              { text: 'Microsoft Word 가이드', link: '/ko/guide/word' },
              { text: 'Microsoft PowerPoint 가이드', link: '/ko/guide/powerpoint' }
            ]
          },
          {
            text: '상세 레퍼런스',
            items: [
              { text: '엑셀 함수 & 옵션 규격', link: '/ko/reference/formula' }
            ]
          },
          {
            text: '고객 지원',
            items: [
              { text: '문제 해결 및 FAQ', link: '/ko/guide/faq' }
            ]
          }
        ],
        editLink: {
          pattern: 'https://github.com/naramdash/func_SMILES2IMG/edit/dev/docs/:path',
          text: 'GitHub에서 이 문서 편집'
        },
        footer: {
          message: 'MIT 라이선스에 따라 제공됩니다.',
          copyright: 'Copyright © 2026 KimJuho. Powered by EPAM Indigo, Ketcher & Mol*.'
        }
      }
    }
  },

  themeConfig: {
    socialLinks: [
      { icon: 'github', link: 'https://github.com/naramdash/func_SMILES2IMG' }
    ],
    search: {
      provider: 'local'
    }
  }
})
