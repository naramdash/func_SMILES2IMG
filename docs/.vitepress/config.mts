import { defineConfig, postcssIsolateStyles } from 'vitepress'

const navApp = (app: 'excel' | 'word' | 'powerpoint', label: string) =>
  `<span class="s2i-nav-app s2i-nav-${app}"><img src="/func_SMILES2IMG/icons/${app}.svg" alt="" width="15" height="15" aria-hidden="true" /><span>${label}</span></span>`

const sidebarJump = (label: string, link: string) => ({
  text: `<span class="s2i-sb-jump"><span>${label}</span><svg class="s2i-sb-jump-icon" viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg></span>`,
  docFooterText: label,
  link
})

const sidebarIconJump = (iconFile: string, label: string, link: string) => ({
  text: `<span class="s2i-sb-jump s2i-sb-app"><img src="/func_SMILES2IMG/icons/${iconFile}" alt="" width="15" height="15" aria-hidden="true" /><span>${label}</span><svg class="s2i-sb-jump-icon" viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg></span>`,
  docFooterText: label,
  link
})

const sidebarCategoryJump = (label: string, link: string) => ({
  text: `<span class="s2i-sb-jump s2i-sb-cat-jump"><span>${label}</span><svg class="s2i-sb-jump-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg></span>`,
  docFooterText: label,
  link,
  items: []
})

const sidebarGroupIcon = (iconFile: string, label: string) =>
  `<span class="s2i-sb-group-icon"><img src="/func_SMILES2IMG/icons/${iconFile}" alt="" width="15" height="15" aria-hidden="true" /><span>${label}</span></span>`

const sidebarMathF = (label: string) =>
  `<math><mi>f</mi></math>: ${label}`

const enGuideItems = [
  { text: 'Introduction', link: '/guide/' },
  { text: 'Installation', link: '/guide/install' },
  { text: 'Getting Started', link: '/guide/getting-started' },
  { text: 'Playbooks', link: '/guide/playbooks' }
]

const enGuideOpen = { text: 'Guide', collapsed: false, items: enGuideItems }
const enInstallLink = sidebarCategoryJump('Install', '/guide/install')

const koGuideItems = [
  { text: '소개', link: '/ko/guide/' },
  { text: '설치', link: '/ko/guide/install' },
  { text: '시작하기', link: '/ko/guide/getting-started' },
  { text: '예제 파일', link: '/ko/guide/playbooks' }
]

const koGuideOpen = { text: '가이드', collapsed: false, items: koGuideItems }
const koInstallLink = sidebarCategoryJump('설치', '/ko/guide/install')

export default defineConfig({
  title: 'func_SMILES2IMG',
  description: 'Free & Open-Source Chemical Structure & 3D Protein Add-in Suite for Microsoft Excel, Word, and PowerPoint (func_SMILES2IMG)',
  base: '/func_SMILES2IMG/',
  cleanUrls: true,
  lastUpdated: true,

  head: [
    ['link', { rel: 'icon', type: 'image/svg+xml', href: '/func_SMILES2IMG/favicon.svg' }],
    ['link', { rel: 'alternate icon', type: 'image/png', sizes: '64x64', href: '/func_SMILES2IMG/favicon.png' }]
  ],

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
            text: 'v1.9.1',
            items: [
              {
                text: 'Release Notes (v1.9.1)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.9.1'
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
                sidebarIconJump('excel.svg', 'Excel', '/excel/'),
                sidebarIconJump('word.svg', 'Word', '/word/'),
                sidebarIconJump('powerpoint.svg', 'PowerPoint', '/powerpoint/')
              ]
            },
            {
              text: 'Editor',
              collapsed: false,
              items: [
                sidebarIconJump('ketcher.png', 'Ketcher (2D)', '/editor/'),
                sidebarIconJump('molstar.png', 'Mol* (3D)', '/editor/molstar'),
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
            enInstallLink,
            {
              text: 'Getting Started',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/excel/' },
                { text: 'Playbook (.xlsx)', link: '/excel/playbook' },
                { text: 'Use Cases', link: '/excel/use-cases' },
                { text: 'All 49 Functions', link: '/excel/formulas' }
              ]
            },
            {
              text: sidebarMathF('Structure Images'),
              collapsed: false,
              items: [
                { text: 'SMILES, InChI', link: '/excel/structures' },
                { text: 'Names, Solvents, CAS', link: '/excel/name-cas' },
                { text: 'Reaction Schemes (KET)', link: '/excel/reactions' },
                { text: 'Styles & Highlights', link: '/excel/styles-reactions' },
                { text: 'Legacy Excel (.FLOAT)', link: '/excel/legacy-float' }
              ]
            },
            {
              text: sidebarMathF('Properties & 3D'),
              collapsed: false,
              items: [
                { text: 'Formula, MW, Exact Mass', link: '/excel/properties' },
                { text: 'LogP, TPSA & 3D', link: '/excel/lipinski-adme' }
              ]
            },
            {
              text: sidebarMathF('SAR & Identifiers'),
              collapsed: false,
              items: [
                { text: 'Substructure Search', link: '/excel/sar' },
                { text: 'Identifier Conversion', link: '/excel/data-sdf' }
              ]
            },
            {
              text: 'Sheet Tools & 3D',
              collapsed: false,
              items: [
                { text: 'SDF Import', link: '/excel/sdf' },
                { text: 'Batch Style, TSV, 3D', link: '/excel/tools-export' },
                { text: 'Ribbon & Context Menus', link: '/excel/ribbon' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Formula Errors', link: '/excel/troubleshooting/formula-errors' },
                { text: 'Image Sizing & .FLOAT', link: '/excel/troubleshooting/image-sizing' }
              ]
            }
          ],
          '/word/': [
            enInstallLink,
            {
              text: 'Getting Started',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/word/' },
                { text: 'Playbook (.docx)', link: '/word/playbook' },
                { text: 'Use Cases', link: '/word/use-cases' }
              ]
            },
            {
              text: 'Manuscript Workflows',
              collapsed: false,
              items: [
                { text: '1. 2D Structures & Reaction Schemes', link: '/word/structures' },
                { text: '2. SAR Tables & Experimental Properties', link: '/word/tables-properties' },
                { text: '3. 3D Protein Figures (Mol*)', link: '/word/3d-figures' },
                { text: '4. Journal Finish, SVG, Sharing', link: '/word/publishing' }
              ]
            },
            {
              text: 'Reference',
              collapsed: false,
              items: [
                { text: 'Ribbon & Context Menus', link: '/word/reference' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Double-Click & Editing', link: '/word/troubleshooting/double-click' },
                { text: 'Table Clipping & Batch Style', link: '/word/troubleshooting/table-layout' }
              ]
            }
          ],
          '/powerpoint/': [
            enInstallLink,
            {
              text: 'Getting Started',
              collapsed: false,
              items: [
                { text: 'Overview', link: '/powerpoint/' },
                { text: 'Playbook (.pptx)', link: '/powerpoint/playbook' },
                { text: 'Use Cases', link: '/powerpoint/use-cases' }
              ]
            },
            {
              text: 'Slide & Poster Workflows',
              collapsed: false,
              items: [
                { text: '1. 2D Structures & Reaction Slides', link: '/powerpoint/slides-structures' },
                { text: '2. Dark Mode, Cards, Deck Restyle', link: '/powerpoint/dark-mode' },
                { text: '3. 360° 3D Protein Slides (Mol* & Morph)', link: '/powerpoint/3d-proteins' },
                { text: '4. A0 Posters (SVG), Q&A, Podium Sharing', link: '/powerpoint/publishing-qa' }
              ]
            },
            {
              text: 'Reference',
              collapsed: false,
              items: [
                { text: 'Ribbon & Context Menus', link: '/powerpoint/reference' }
              ]
            },
            {
              text: 'Troubleshooting',
              collapsed: true,
              items: [
                { text: 'Dark Slides & Grouped Shapes', link: '/powerpoint/troubleshooting/dark-slides' },
                { text: '3D Models & Double-Click', link: '/powerpoint/troubleshooting/3d-editing' }
              ]
            }
          ],
          '/editor/': [
            enInstallLink,
            {
              text: sidebarGroupIcon('ketcher.png', 'Ketcher (2D)'),
              collapsed: false,
              items: [
                { text: '2D Sketcher & Reaction Schemes', link: '/editor/' },
                { text: 'Ketcher Official Website', link: 'https://lifescience.opensource.epam.com/ketcher/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher User Manual', link: 'https://github.com/epam/ketcher/blob/master/documentation/help.md', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher GitHub', link: 'https://github.com/epam/ketcher', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: sidebarGroupIcon('molstar.png', 'Mol* (3D)'),
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
                { text: 'ACS B&W Batch Style, SVG, Property Tables', link: '/editor/publishing' }
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
            text: 'v1.9.1',
            items: [
              {
                text: '최신 릴리스 노트 (v1.9.1)',
                link: 'https://github.com/naramdash/func_SMILES2IMG/releases/tag/v1.9.1'
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
                sidebarIconJump('excel.svg', 'Excel', '/ko/excel/'),
                sidebarIconJump('word.svg', 'Word', '/ko/word/'),
                sidebarIconJump('powerpoint.svg', 'PowerPoint', '/ko/powerpoint/')
              ]
            },
            {
              text: 'Editor',
              collapsed: false,
              items: [
                sidebarIconJump('ketcher.png', 'Ketcher (2D)', '/ko/editor/'),
                sidebarIconJump('molstar.png', 'Mol* (3D)', '/ko/editor/molstar'),
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
            koInstallLink,
            {
              text: '시작하기',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/excel/' },
                { text: '예제 파일 (.xlsx)', link: '/ko/excel/playbook' },
                { text: '활용 사례', link: '/ko/excel/use-cases' },
                { text: '49개 함수 전체 개요', link: '/ko/excel/formulas' }
              ]
            },
            {
              text: sidebarMathF('구조식 이미지'),
              collapsed: false,
              items: [
                { text: 'SMILES · InChI', link: '/ko/excel/structures' },
                { text: '시약명 · 용매 · CAS', link: '/ko/excel/name-cas' },
                { text: '합성 반응식 (KET)', link: '/ko/excel/reactions' },
                { text: '스타일 · 하이라이트', link: '/ko/excel/styles-reactions' },
                { text: '구버전 Excel (.FLOAT)', link: '/ko/excel/legacy-float' }
              ]
            },
            {
              text: sidebarMathF('물성 · 3D 지표'),
              collapsed: false,
              items: [
                { text: '화학식 · 분자량 · 정밀질량', link: '/ko/excel/properties' },
                { text: 'LogP · TPSA · 3D', link: '/ko/excel/lipinski-adme' }
              ]
            },
            {
              text: sidebarMathF('SAR · 식별자'),
              collapsed: false,
              items: [
                { text: '부분구조 검색 (SMARTS)', link: '/ko/excel/sar' },
                { text: '표기법 변환 · InChIKey', link: '/ko/excel/data-sdf' }
              ]
            },
            {
              text: '시트 도구 · 3D',
              collapsed: false,
              items: [
                { text: 'SDF 파일 가져오기', link: '/ko/excel/sdf' },
                { text: '일괄 스타일 · TSV · 3D', link: '/ko/excel/tools-export' },
                { text: '리본 · 우클릭 메뉴', link: '/ko/excel/ribbon' }
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
            koInstallLink,
            {
              text: '시작하기',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/word/' },
                { text: '예제 파일 (.docx)', link: '/ko/word/playbook' },
                { text: '활용 사례', link: '/ko/word/use-cases' }
              ]
            },
            {
              text: '논문 · 보고서 실전 가이드',
              collapsed: false,
              items: [
                { text: '1. 2D 구조식 · 반응식 작성', link: '/ko/word/structures' },
                { text: '2. 화합물 표(SAR) & 실험 물성(SI)', link: '/ko/word/tables-properties' },
                { text: '3. 3D 단백질 Figure (Mol*)', link: '/ko/word/3d-figures' },
                { text: '4. 투고 마감 · 벡터 SVG · 공동저자 공유', link: '/ko/word/publishing' }
              ]
            },
            {
              text: '레퍼런스',
              collapsed: false,
              items: [
                { text: '리본 · 우클릭 메뉴', link: '/ko/word/reference' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '더블클릭 · 재편집', link: '/ko/word/troubleshooting/double-click' },
                { text: '표 잘림 · 일괄 스타일', link: '/ko/word/troubleshooting/table-layout' }
              ]
            }
          ],
          '/ko/powerpoint/': [
            koInstallLink,
            {
              text: '시작하기',
              collapsed: false,
              items: [
                { text: '개요', link: '/ko/powerpoint/' },
                { text: '예제 파일 (.pptx)', link: '/ko/powerpoint/playbook' },
                { text: '활용 사례', link: '/ko/powerpoint/use-cases' }
              ]
            },
            {
              text: '발표 · 포스터 실전 가이드',
              collapsed: false,
              items: [
                { text: '1. 2D 구조식 · 반응식 슬라이드', link: '/ko/powerpoint/slides-structures' },
                { text: '2. 다크 모드 · 후보 카드 · 덱 테마 전환', link: '/ko/powerpoint/dark-mode' },
                { text: '3. 360° 3D 단백질 발표 (Mol* & 모핑)', link: '/ko/powerpoint/3d-proteins' },
                { text: '4. A0 포스터(SVG) · Q&A · 발표장 공유', link: '/ko/powerpoint/publishing-qa' }
              ]
            },
            {
              text: '레퍼런스',
              collapsed: false,
              items: [
                { text: '리본 · 우클릭 메뉴', link: '/ko/powerpoint/reference' }
              ]
            },
            {
              text: '트러블슈팅',
              collapsed: true,
              items: [
                { text: '다크 슬라이드 · 그룹 도형', link: '/ko/powerpoint/troubleshooting/dark-slides' },
                { text: '3D 모델 · 더블클릭 편집', link: '/ko/powerpoint/troubleshooting/3d-editing' }
              ]
            }
          ],
          '/ko/editor/': [
            koInstallLink,
            {
              text: sidebarGroupIcon('ketcher.png', 'Ketcher (2D)'),
              collapsed: false,
              items: [
                { text: '2D 구조·반응식 편집기', link: '/ko/editor/' },
                { text: 'Ketcher 공식 웹사이트', link: 'https://lifescience.opensource.epam.com/ketcher/', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher 공식 사용자 가이드', link: 'https://github.com/epam/ketcher/blob/master/documentation/help.md', target: '_blank', rel: 'noopener noreferrer' },
                { text: 'Ketcher GitHub', link: 'https://github.com/epam/ketcher', target: '_blank', rel: 'noopener noreferrer' }
              ]
            },
            {
              text: sidebarGroupIcon('molstar.png', 'Mol* (3D)'),
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
    logo: {
      light: '/logo-light.svg',
      dark: '/logo-dark.svg',
      alt: 'func_SMILES2IMG'
    },
    socialLinks: [
      { icon: 'github', link: 'https://github.com/naramdash/func_SMILES2IMG' }
    ],
    search: {
      provider: 'local'
    }
  },

  vite: {
    css: {
      postcss: {
        plugins: [
          postcssIsolateStyles({
            includeFiles: [/vp-doc\.css/, /base\.css/]
          })
        ]
      }
    },
    server: {
      allowedHosts: true
    }
  }
})
