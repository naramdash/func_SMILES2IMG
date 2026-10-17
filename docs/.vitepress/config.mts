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
        outline: {
          level: [2, 3],
          label: 'On this page'
        },
        nav: [
          { text: 'Quickstart', link: '/guide/getting-started', activeMatch: '^/guide/' },
          { text: 'Excel Guide', link: '/guide/excel' },
          { text: 'Formula Cheat Sheet', link: '/reference/formula', activeMatch: '^/reference/' },
          { text: 'Formula Generator', link: '/tools/formula-builder', activeMatch: '^/tools/' },
          { text: 'v1.7.0', link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.0' }
        ],
        sidebar: [
          {
            text: 'Getting Started',
            items: [
              { text: '🚀 Install & 1-Min Quickstart', link: '/guide/getting-started' }
            ]
          },
          {
            text: 'Microsoft Excel Guides',
            items: [
              { text: '📗 Excel Basics: Tables & Styling', link: '/guide/excel' },
              { text: '🔬 Excel SAR, Lipinski & SDF Files', link: '/guide/excel-cheminformatics' }
            ]
          },
          {
            text: 'Word, PowerPoint & Publishing',
            items: [
              { text: '📘 Word: Manuscripts & Tables', link: '/guide/word' },
              { text: '📙 PowerPoint: Slides & Dark Mode', link: '/guide/powerpoint' },
              { text: '🎨 Batch Style, Vector SVG & Export', link: '/guide/publishing' }
            ]
          },
          {
            text: '2D Sketcher & 3D Viewer',
            items: [
              { text: '✏️ Ketcher 2D Sketcher', link: '/guide/ketcher' },
              { text: '🧬 Mol* 3D Protein Viewer', link: '/guide/molstar' }
            ]
          },
          {
            text: 'Reference & Help',
            items: [
              { text: '📐 Excel Formula Cheat Sheet', link: '/reference/formula' },
              { text: '❓ Troubleshooting & FAQ', link: '/guide/faq' }
            ]
          },
          {
            text: 'Interactive Web Tools',
            items: [
              { text: '🛠️ Formula Generator & Style Simulator', link: '/tools/formula-builder' }
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
      description: 'Microsoft Excel, Word, PowerPoint용 무료 오픈소스 화학 구조식 & 3D 단백질 추가 기능',
      themeConfig: {
        outline: {
          level: [2, 3],
          label: '이 페이지 목차'
        },
        nav: [
          { text: '빠른 시작', link: '/ko/guide/getting-started', activeMatch: '^/ko/guide/' },
          { text: 'Excel 가이드', link: '/ko/guide/excel' },
          { text: '함수 치트시트', link: '/ko/reference/formula', activeMatch: '^/ko/reference/' },
          { text: '수식 표현 생성기', link: '/ko/tools/formula-builder', activeMatch: '^/ko/tools/' },
          { text: 'v1.7.0', link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.0' }
        ],
        sidebar: [
          {
            text: '시작하기',
            items: [
              { text: '🚀 설치 및 1분 첫걸음', link: '/ko/guide/getting-started' }
            ]
          },
          {
            text: 'Microsoft Excel 가이드',
            items: [
              { text: '📗 Excel 기초: 구조식 표 & 스타일', link: '/ko/guide/excel' },
              { text: '🔬 Excel 심화: 물성 · SAR · SDF', link: '/ko/guide/excel-cheminformatics' }
            ]
          },
          {
            text: 'Word · PowerPoint & 출판 도구',
            items: [
              { text: '📘 Word: 학술 논문 & 보고서 작성', link: '/ko/guide/word' },
              { text: '📙 PowerPoint: 학회 발표 슬라이드', link: '/ko/guide/powerpoint' },
              { text: '🎨 출판 스타일 · SVG 벡터 · 물성 복사', link: '/ko/guide/publishing' }
            ]
          },
          {
            text: '2D 스케처 & 3D 뷰어',
            items: [
              { text: '✏️ Ketcher 2D 구조/반응식 편집기', link: '/ko/guide/ketcher' },
              { text: '🧬 Mol* 3D 단백질 뷰어', link: '/ko/guide/molstar' }
            ]
          },
          {
            text: '수식 사전 & 문제 해결',
            items: [
              { text: '📐 전체 엑셀 함수 치트시트', link: '/ko/reference/formula' },
              { text: '❓ 자주 묻는 질문 (FAQ)', link: '/ko/guide/faq' }
            ]
          },
          {
            text: '관련 툴 (Interactive Tools)',
            items: [
              { text: '🛠️ 엑셀 함수 표현 생성기 & 스타일 프리뷰', link: '/ko/tools/formula-builder' }
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
