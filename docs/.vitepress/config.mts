import { defineConfig } from 'vitepress'

export default defineConfig({
  title: 'func_SMILES2IMG',
  description: 'Free & Open-Source Chemical Structure & 3D Protein Add-in Suite for Microsoft Excel, Word, and PowerPoint (func_SMILES2IMG)',
  base: '/func_SMILES2IMG/',
  cleanUrls: true,
  lastUpdated: true,

  head: [],

  locales: {
    root: {
      label: 'English',
      lang: 'en',
      title: 'func_SMILES2IMG',
      description: 'Free & Open-Source Chemical Structure & 3D Protein Add-in Suite for Microsoft Excel, Word, and PowerPoint (func_SMILES2IMG)',
      themeConfig: {
        outline: {
          level: [2, 3],
          label: 'On this page'
        },
        docFooter: {
          prev: 'Previous Step',
          next: 'Next Step'
        },
        nav: [
          {
            text: 'v1.8.0',
            items: [
              {
                text: 'Release Notes (v1.8.0)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.8.0'
              },
              {
                text: 'All Releases',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases'
              },
              {
                text: 'Issue Tracker & Feedback',
                link: 'https://github.com/naramdash/func_SMILES2IMG/issues'
              }
            ]
          }
        ],
        sidebar: [
          {
            text: '📦 Getting Started',
            items: [
              { text: '🚀 1-Min Install & First Molecule', link: '/guide/getting-started' }
            ]
          },
          {
            text: '📗 Microsoft Excel (Recipe Cookbook)',
            items: [
              { text: '📘 Recipes 1–5: Tables, Name/CAS, Styles & Reactions', link: '/guide/excel' },
              { text: '🔬 Recipes 6–10: MW, Lipinski, SAR & SDF Files', link: '/guide/excel-cheminformatics' },
              { text: '⚡ Interactive Formula Builder', link: '/tools/formula-builder' },
              { text: '📑 40 Formula & Style Cheat Sheet', link: '/reference/formula' }
            ]
          },
          {
            text: '📘 Microsoft Word (Manuscript Tutorial)',
            items: [
              { text: '📝 6-Chapter Manuscript Tutorial & Task Index', link: '/guide/word' },
              { text: '📐 ACS B&W Batch Style, SVG & Property Tables', link: '/guide/publishing' }
            ]
          },
          {
            text: '📙 Microsoft PowerPoint (Slide Gallery)',
            items: [
              { text: '🎨 Visual Slide Gallery & 3D Proteins', link: '/guide/powerpoint' }
            ]
          },
          {
            text: '🧩 Shared Editors & Support',
            items: [
              { text: '✏️ Ketcher 2D Sketcher (All Apps)', link: '/guide/ketcher' },
              { text: '🧬 Mol* 3D Protein Viewer (All Apps)', link: '/guide/molstar' },
              { text: '❓ Troubleshooting & #NAME? Fix (FAQ)', link: '/guide/faq' }
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
      description: 'Microsoft Excel, Word, PowerPoint용 무료 오픈소스 화학 구조식 & 3D 단백질 추가 기능 (func_SMILES2IMG)',
      themeConfig: {
        outline: {
          level: [2, 3],
          label: '이 페이지 목차'
        },
        docFooter: {
          prev: '이전 단계 가이드',
          next: '다음 단계 가이드'
        },
        nav: [
          {
            text: 'v1.8.0',
            items: [
              {
                text: '최신 릴리스 노트 (v1.8.0)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.8.0'
              },
              {
                text: '전체 버전 릴리스 목록',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases'
              },
              {
                text: '버그 제보 및 기능 제안 (Issues)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/issues'
              }
            ]
          }
        ],
        sidebar: [
          {
            text: '📦 시작하기 (Getting Started)',
            items: [
              { text: '🚀 1분 설치 및 첫 분자 그리기', link: '/ko/guide/getting-started' }
            ]
          },
          {
            text: '📗 Microsoft Excel (실무 레시피 & 40개 함수)',
            items: [
              { text: '📘 레시피 1~5: 구조식 표 · 이름/CAS · 스타일 · 반응식', link: '/ko/guide/excel' },
              { text: '🔬 레시피 6~10: 물성 · Lipinski · SAR & SDF 연동', link: '/ko/guide/excel-cheminformatics' },
              { text: '⚡ 클릭으로 만드는 수식 생성기', link: '/ko/tools/formula-builder' },
              { text: '📑 40개 전체 수식 & 스타일 키워드 사전', link: '/ko/reference/formula' }
            ]
          },
          {
            text: '📘 Microsoft Word (논문 1편 완성 튜토리얼)',
            items: [
              { text: '📝 6장 연속 논문 원고 튜토리얼 & 빠른 찾아보기', link: '/ko/guide/word' },
              { text: '📐 투고용 ACS 흑백 변환 · SVG · 물성 표 복사', link: '/ko/guide/publishing' }
            ]
          },
          {
            text: '📙 Microsoft PowerPoint (슬라이드 갤러리)',
            items: [
              { text: '🎨 비주얼 슬라이드 갤러리 & 3D 단백질 발표', link: '/ko/guide/powerpoint' }
            ]
          },
          {
            text: '🧩 공통 에디터 & 문제 해결',
            items: [
              { text: '✏️ Ketcher 2D 구조·반응식 편집기 (공통)', link: '/ko/guide/ketcher' },
              { text: '🧬 Mol* 3D 단백질 뷰어 (공통)', link: '/ko/guide/molstar' },
              { text: '❓ 설치 및 #NAME? 문제 해결 (FAQ)', link: '/ko/guide/faq' }
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
