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
            text: 'v1.7.1',
            items: [
              {
                text: 'Release Notes (v1.7.1)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.1'
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
            text: '📊 By Office App (Workflows)',
            items: [
              { text: '📗 Excel Part 1: Structure Tables & Name/CAS', link: '/guide/excel' },
              { text: '🔬 Excel Part 2: MW, Lipinski, SAR & SDF', link: '/guide/excel-cheminformatics' },
              { text: '📘 Word: Manuscripts & Summary Tables', link: '/guide/word' },
              { text: '📙 PowerPoint: Dark Mode & 3D Slides', link: '/guide/powerpoint' }
            ]
          },
          {
            text: '🎨 Core Visual Tools (All Apps)',
            items: [
              { text: '✏️ Draw 2D Structures & Reactions (Ketcher)', link: '/guide/ketcher' },
              { text: '🧬 Insert & Rotate 3D Proteins (Mol*)', link: '/guide/molstar' },
              { text: '📐 1-Click ACS B&W & Illustrator SVG', link: '/guide/publishing' }
            ]
          },
          {
            text: '🛠️ Interactive Web Tools',
            items: [
              { text: '⚡ Interactive Formula Builder', link: '/tools/formula-builder' }
            ]
          },
          {
            text: '📑 Reference & Support',
            items: [
              { text: '📑 25+ Formula & Style Cheat Sheet', link: '/reference/formula' },
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
            text: 'v1.7.1',
            items: [
              {
                text: '최신 릴리스 노트 (v1.7.1)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.7.1'
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
            text: '📊 앱별 실무 가이드 (By Office App)',
            items: [
              { text: '📗 Excel 1부: 구조식 테이블 & 이름/CAS 변환', link: '/ko/guide/excel' },
              { text: '🔬 Excel 2부: 물성(MW·Lipinski), SAR & SDF 연동', link: '/ko/guide/excel-cheminformatics' },
              { text: '📘 Word: 논문·보고서 본문 및 표에 구조 넣기', link: '/ko/guide/word' },
              { text: '📙 PowerPoint: 다크 모드 & 3D 발표 슬라이드', link: '/ko/guide/powerpoint' }
            ]
          },
          {
            text: '🎨 공통 핵심 기능 (Core Visual Tools)',
            items: [
              { text: '✏️ 2D 구조 및 반응식 직접 그리기 (Ketcher)', link: '/ko/guide/ketcher' },
              { text: '🧬 3D 단백질(PDB) 삽입 및 360° 회전 (Mol*)', link: '/ko/guide/molstar' },
              { text: '📐 논문 투고용 ACS 일괄 변환 & SVG 내보내기', link: '/ko/guide/publishing' }
            ]
          },
          {
            text: '🛠️ 웹 인터랙티브 도구',
            items: [
              { text: '⚡ 클릭으로 만드는 수식 생성기', link: '/ko/tools/formula-builder' }
            ]
          },
          {
            text: '📑 레퍼런스 및 고객 지원',
            items: [
              { text: '📑 25+ 전체 수식 & 스타일 키워드 사전', link: '/ko/reference/formula' },
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
