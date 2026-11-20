import { defineConfig } from 'vitepress'

const navApp = (app: 'excel' | 'word' | 'powerpoint', label: string) =>
  `<span class="s2i-nav-app s2i-nav-${app}"><img src="/func_SMILES2IMG/icons/${app}.svg" alt="" width="15" height="15" aria-hidden="true" /><span>${label}</span></span>`

const sidebarJump = (label: string, link: string) => ({
  text: `<span class="s2i-sb-jump"><span>${label}</span><svg class="s2i-sb-jump-icon" viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg></span>`,
  docFooterText: label,
  link
})

const enGuideItems = [
  { text: 'Introduction', link: '/guide/' },
  { text: 'Installation', link: '/guide/install' },
  { text: 'Getting Started', link: '/guide/getting-started' },
  { text: 'Playbooks', link: '/guide/playbooks' }
]

const enGuideOpen = { text: 'Guide', collapsed: false, items: enGuideItems }
const enGuideClosed = { text: 'Guide', collapsed: true, items: enGuideItems }

const koGuideItems = [
  { text: '소개', link: '/ko/guide/' },
  { text: '설치', link: '/ko/guide/install' },
  { text: '시작하기', link: '/ko/guide/getting-started' },
  { text: '예제 파일', link: '/ko/guide/playbooks' }
]

const koGuideOpen = { text: '가이드', collapsed: false, items: koGuideItems }
const koGuideClosed = { text: '가이드', collapsed: true, items: koGuideItems }

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
          prev: 'Previous Page',
          next: 'Next Page'
        },
        nav: [
          { text: 'Guide', link: '/guide/', activeMatch: '^/guide/' },
          { text: navApp('excel', 'Excel'), link: '/excel/', activeMatch: '^/excel/' },
          { text: navApp('word', 'Word'), link: '/word/', activeMatch: '^/word/' },
          { text: navApp('powerpoint', 'PowerPoint'), link: '/powerpoint/', activeMatch: '^/powerpoint/' },
          { text: 'Editor', link: '/editor/', activeMatch: '^/editor/' },
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
        sidebar: {
          '/guide/': [
            enGuideOpen,
            {
              text: 'Office Apps',
              collapsed: false,
              items: [
                sidebarJump('Microsoft Excel', '/excel/'),
                sidebarJump('Microsoft Word', '/word/'),
                sidebarJump('Microsoft PowerPoint', '/powerpoint/')
              ]
            },
            {
              text: 'Editor',
              collapsed: false,
              items: [
                sidebarJump('Ketcher (2D)', '/editor/'),
                sidebarJump('Mol* (3D)', '/editor/molstar'),
                sidebarJump('Style & Export', '/editor/publishing')
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Installation Blocks', link: '/guide/troubleshooting/install' },
                { text: 'Missing Ribbon Tab', link: '/guide/troubleshooting/ribbon-tab' },
                { text: 'Document Sharing', link: '/guide/troubleshooting/document-sharing' },
                { text: 'FAQ', link: '/guide/faq' }
              ]
            }
          ],
          '/excel/': [
            enGuideClosed,
            {
              text: 'Rendering & Drawing',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/excel/' },
                { text: 'Structure Tables & Name/CAS', link: '/excel/structures' },
                { text: 'Styles, Highlighting & Reactions', link: '/excel/styles-reactions' },
                { text: 'Legacy Excel (.FLOAT)', link: '/excel/legacy-float' }
              ]
            },
            {
              text: 'Cheminformatics & Data',
              collapsed: false,
              items: [
                { text: 'Properties & Lipinski Ro5', link: '/excel/properties' },
                { text: 'Substructure & Similarity (SAR)', link: '/excel/sar' },
                { text: 'Notation Conversion', link: '/excel/conversions' },
                { text: 'SDF Import & Export', link: '/excel/sdf' }
              ]
            },
            {
              text: 'Reference & Tools',
              collapsed: false,
              items: [
                { text: 'Ribbon & Context Menu', link: '/excel/ribbon' },
                { text: '40 Formulas Reference', link: '/excel/formulas' },
                { text: 'Formula Builder', link: '/excel/formula-builder' },
                { text: 'Playbook (.xlsx)', link: '/excel/playbook' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Formula Errors', link: '/excel/troubleshooting/formula-errors' },
                { text: 'Image Sizing', link: '/excel/troubleshooting/image-sizing' }
              ]
            }
          ],
          '/word/': [
            enGuideClosed,
            {
              text: 'Creating Structures',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/word/' },
                { text: 'Insert & Paste Structures', link: '/word/insert' },
                { text: 'Compound Tables & Properties', link: '/word/tables' },
                { text: 'Reaction Schemes', link: '/word/reactions' }
              ]
            },
            {
              text: 'Editing & Publishing',
              collapsed: false,
              items: [
                { text: 'In-Place Editing & 3D Proteins', link: '/word/editing-3d' },
                { text: 'ACS B&W & Vector SVG', link: '/word/publishing' },
                { text: 'Ribbon & Context Menu', link: '/word/ribbon' },
                { text: 'Playbook (.docx)', link: '/word/playbook' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Double-Click Editing', link: '/word/troubleshooting/double-click' },
                { text: 'Table Layout', link: '/word/troubleshooting/table-layout' }
              ]
            }
          ],
          '/powerpoint/': [
            enGuideClosed,
            {
              text: 'Slide Workflows',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/powerpoint/' },
                { text: 'Dark Mode & Screening Cards', link: '/powerpoint/dark-mode' },
                { text: 'TextBox Quick Convert', link: '/powerpoint/quick-convert' },
                { text: 'Reaction Mechanism Slides', link: '/powerpoint/reactions' }
              ]
            },
            {
              text: '3D, Q&A & Styling',
              collapsed: false,
              items: [
                { text: '3D Protein Slides (Mol*)', link: '/powerpoint/3d-proteins' },
                { text: 'Live Q&A Molecule Info', link: '/powerpoint/molecule-info' },
                { text: 'Deck-Wide Batch Style & SVG', link: '/powerpoint/batch-style' },
                { text: 'Playbook (.pptx)', link: '/powerpoint/playbook' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Dark Slide Backgrounds', link: '/powerpoint/troubleshooting/dark-slides' },
                { text: '3D Models', link: '/powerpoint/troubleshooting/3d-models' }
              ]
            }
          ],
          '/editor/': [
            enGuideClosed,
            {
              text: 'Ketcher (2D)',
              collapsed: false,
              items: [
                { text: '2D Sketcher & Reaction Schemes', link: '/editor/' },
                { text: 'Ketcher Official Website', link: 'https://lifescience.opensource.epam.com/ketcher/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher User Manual', link: 'https://github.com/epam/ketcher/blob/master/documentation/help.md', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher GitHub', link: 'https://github.com/epam/ketcher', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: 'Mol* (3D)',
              collapsed: false,
              items: [
                { text: '3D Protein Viewer & GLB Models', link: '/editor/molstar' },
                { text: 'Mol* Official Website', link: 'https://molstar.org/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Mol* Viewer Guide', link: 'https://molstar.org/viewer-docs/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'RCSB PDB Search', link: 'https://www.rcsb.org/', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: 'Style & Export',
              collapsed: false,
              items: [
                { text: 'ACS B&W Batch Style, SVG & Property Tables', link: '/editor/publishing' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'WebView2 Loading', link: '/editor/troubleshooting/webview2-loading' },
                { text: 'Offline PDB Files', link: '/editor/troubleshooting/offline-pdb' }
              ]
            }
          ]
        },
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
          prev: '이전 문서',
          next: '다음 문서'
        },
        nav: [
          { text: '가이드', link: '/ko/guide/', activeMatch: '^/ko/guide/' },
          { text: navApp('excel', 'Excel'), link: '/ko/excel/', activeMatch: '^/ko/excel/' },
          { text: navApp('word', 'Word'), link: '/ko/word/', activeMatch: '^/ko/word/' },
          { text: navApp('powerpoint', 'PowerPoint'), link: '/ko/powerpoint/', activeMatch: '^/ko/powerpoint/' },
          { text: 'Editor', link: '/ko/editor/', activeMatch: '^/ko/editor/' },
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
        sidebar: {
          '/ko/guide/': [
            koGuideOpen,
            {
              text: 'Office 앱',
              collapsed: false,
              items: [
                sidebarJump('Microsoft Excel', '/ko/excel/'),
                sidebarJump('Microsoft Word', '/ko/word/'),
                sidebarJump('Microsoft PowerPoint', '/ko/powerpoint/')
              ]
            },
            {
              text: 'Editor',
              collapsed: false,
              items: [
                sidebarJump('Ketcher (2D)', '/ko/editor/'),
                sidebarJump('Mol* (3D)', '/ko/editor/molstar'),
                sidebarJump('스타일 & 출력', '/ko/editor/publishing')
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '설치 보안 차단', link: '/ko/guide/troubleshooting/install' },
                { text: '리본 탭 미표시', link: '/ko/guide/troubleshooting/ribbon-tab' },
                { text: '문서 공유', link: '/ko/guide/troubleshooting/document-sharing' },
                { text: 'FAQ', link: '/ko/guide/faq' }
              ]
            }
          ],
          '/ko/excel/': [
            koGuideClosed,
            {
              text: '구조식 렌더링',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/excel/' },
                { text: '구조식 표 · 이름/CAS', link: '/ko/excel/structures' },
                { text: '스타일 · 하이라이트 · 반응식', link: '/ko/excel/styles-reactions' },
                { text: '구버전 Excel (.FLOAT)', link: '/ko/excel/legacy-float' }
              ]
            },
            {
              text: '물성 · SAR · 데이터',
              collapsed: false,
              items: [
                { text: '물성 · Lipinski 약물성', link: '/ko/excel/properties' },
                { text: '부분구조 · 유사도 (SAR)', link: '/ko/excel/sar' },
                { text: '표기법 변환 (InChI · CAS)', link: '/ko/excel/conversions' },
                { text: 'SDF 파일 가져오기 · 내보내기', link: '/ko/excel/sdf' }
              ]
            },
            {
              text: '레퍼런스 · 도구',
              collapsed: false,
              items: [
                { text: '리본 · 우클릭 메뉴', link: '/ko/excel/ribbon' },
                { text: '40개 함수 사전', link: '/ko/excel/formulas' },
                { text: '수식 생성기', link: '/ko/excel/formula-builder' },
                { text: '예제 파일 (.xlsx)', link: '/ko/excel/playbook' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '수식 오류', link: '/ko/excel/troubleshooting/formula-errors' },
                { text: '이미지 크기 · .FLOAT', link: '/ko/excel/troubleshooting/image-sizing' }
              ]
            }
          ],
          '/ko/word/': [
            koGuideClosed,
            {
              text: '구조식 삽입 · 편집',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/word/' },
                { text: '구조식 삽입 · 텍스트 변환', link: '/ko/word/insert' },
                { text: '화합물 표 · 물성 복사', link: '/ko/word/tables' },
                { text: '합성 반응식 (KET)', link: '/ko/word/reactions' }
              ]
            },
            {
              text: '편집 · 논문 투고',
              collapsed: false,
              items: [
                { text: '더블클릭 편집 · 3D 단백질', link: '/ko/word/editing-3d' },
                { text: 'ACS 흑백 변환 · SVG', link: '/ko/word/publishing' },
                { text: '리본 · 우클릭 메뉴', link: '/ko/word/ribbon' },
                { text: '예제 파일 (.docx)', link: '/ko/word/playbook' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '더블클릭 편집', link: '/ko/word/troubleshooting/double-click' },
                { text: '표 레이아웃 · 스타일', link: '/ko/word/troubleshooting/table-layout' }
              ]
            }
          ],
          '/ko/powerpoint/': [
            koGuideClosed,
            {
              text: '슬라이드 제작',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/powerpoint/' },
                { text: '다크 모드 · 후보 물질 카드', link: '/ko/powerpoint/dark-mode' },
                { text: '텍스트 상자 즉시 변환', link: '/ko/powerpoint/quick-convert' },
                { text: '반응 메커니즘 슬라이드', link: '/ko/powerpoint/reactions' }
              ]
            },
            {
              text: '3D · 발표 · 스타일',
              collapsed: false,
              items: [
                { text: '3D 단백질 슬라이드 (Mol*)', link: '/ko/powerpoint/3d-proteins' },
                { text: '발표 Q&A 물성 조회', link: '/ko/powerpoint/molecule-info' },
                { text: '전체 슬라이드 스타일 · SVG', link: '/ko/powerpoint/batch-style' },
                { text: '예제 파일 (.pptx)', link: '/ko/powerpoint/playbook' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '다크 슬라이드 배경', link: '/ko/powerpoint/troubleshooting/dark-slides' },
                { text: '3D 모델 · 편집', link: '/ko/powerpoint/troubleshooting/3d-models' }
              ]
            }
          ],
          '/ko/editor/': [
            koGuideClosed,
            {
              text: 'Ketcher (2D)',
              collapsed: false,
              items: [
                { text: '2D 구조·반응식 편집기', link: '/ko/editor/' },
                { text: 'Ketcher 공식 웹사이트', link: 'https://lifescience.opensource.epam.com/ketcher/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher 공식 사용자 가이드', link: 'https://github.com/epam/ketcher/blob/master/documentation/help.md', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher GitHub', link: 'https://github.com/epam/ketcher', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: 'Mol* (3D)',
              collapsed: false,
              items: [
                { text: '3D 단백질 뷰어 & GLB 모델', link: '/ko/editor/molstar' },
                { text: 'Mol* 공식 웹사이트', link: 'https://molstar.org/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Mol* 뷰어 공식 가이드', link: 'https://molstar.org/viewer-docs/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'RCSB PDB 단백질 검색', link: 'https://www.rcsb.org/', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: '스타일 & 출력',
              collapsed: false,
              items: [
                { text: '투고용 ACS 흑백 변환 · SVG · 물성 표 복사', link: '/ko/editor/publishing' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: 'WebView2 로딩 오류', link: '/ko/editor/troubleshooting/webview2-loading' },
                { text: '오프라인 PDB 파일', link: '/ko/editor/troubleshooting/offline-pdb' }
              ]
            }
          ]
        },
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
  },

  vite: {
    server: {
      allowedHosts: true
    }
  }
})
