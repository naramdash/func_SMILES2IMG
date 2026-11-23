<script setup lang="ts">
import { computed } from 'vue'
import { useData, withBase } from 'vitepress'

const props = defineProps<{
  app: 'word' | 'powerpoint'
}>()

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const content = computed(() => {
  if (props.app === 'word') {
    if (isKo.value) {
      return {
        ctaClass: 'is-word',
        cards: [
          {
            title: '🔒 로컬 우선 실행',
            desc: '2D 렌더링(EPAM Indigo), Ketcher 2D, 반응식 KET 저장, SVG 복사, 12종 물성 계산이 모두 내 PC 로컬에서 실행되어 미공개 연구 구조식이 외부 서버로 나가지 않습니다.'
          },
          {
            title: '🍎 Mac · 공동저자 환경 호환',
            desc: '표준 PNG 그림과 3D <code>.glb</code> 모델로 저장되므로 애드인이 없는 Mac·Windows·웹 Word에서도 그림이 깨지지 않으며, 메타데이터 덕분에 Windows에서 다시 열면 더블클릭 수정도 유지됩니다.'
          },
          {
            title: '⚡ 무료 MIT · 관리자 권한 불필요',
            desc: '무료·오픈소스(MIT)이며 Windows 관리자(Admin) 권한 없이 사용자 영역에 설치됩니다. 필요 없을 때는 한 줄 명령으로 제거할 수 있습니다.'
          }
        ],
        ctaTitle: 'Word · Excel · PowerPoint 통합 설치',
        ctaMeta:
          '열려 있는 오피스 창을 닫고 <strong>Windows PowerShell</strong>에 아래 명령어 한 줄을 붙여넣어 실행한 뒤, Word를 다시 열면 상단에 <code>[SMILES]</code> 탭이 나타납니다:',
        installHref: '/ko/guide/install',
        installLabel: '상세 설치 가이드 & 사내망 오프라인 ZIP 설치 →',
        playbookHref: '/ko/word/playbook',
        playbookLabel: '6페이지 Word 실습 원고(Playbook .docx) 미리보기 →'
      }
    }
    return {
      ctaClass: 'is-word',
      cards: [
        {
          title: '🔒 Local-First Execution',
          desc: '2D rendering (EPAM Indigo), Ketcher 2D, reaction KET storage, SVG export, and all 17 property calculations run offline on your PC—unpublished structures never leave your machine.'
        },
        {
          title: '🍎 Compatible Anywhere You Share',
          desc: 'Structures are saved as standard PNG images and 3D <code>.glb</code> models with Alt Text tags. Co-authors on macOS, reviewers, and Word Online see the same figures without installing the add-in.'
        },
        {
          title: '⚡ Free MIT · No Admin Rights Needed',
          desc: 'Free and open-source (MIT). Installs into your user profile across Excel, Word, and PowerPoint without Windows administrator rights—and uninstalls with a single command.'
        }
      ],
      ctaTitle: 'Install Across Word, Excel & PowerPoint',
      ctaMeta:
        'Close Office apps, paste this command into <strong>Windows PowerShell</strong> (no admin rights required), then reopen Microsoft Word to see the <code>[SMILES]</code> ribbon tab:',
      installHref: '/guide/install',
      installLabel: 'Full Install Guide & Offline Corporate ZIP →',
      playbookHref: '/word/playbook',
      playbookLabel: 'Preview 6-Page Word Playbook (.docx) →'
    }
  }

  if (isKo.value) {
    return {
      ctaClass: 'is-ppt',
      cards: [
        {
          title: '🔒 로컬 우선 실행',
          desc: 'Ketcher 2D 스케치, Indigo 렌더링, SVG 복사, 12종 물성 조회가 모두 내 PC 로컬에서 실행되어 미공개 후보 물질 슬라이드가 외부로 나가지 않습니다.'
        },
        {
          title: '🎤 발표장 PC · Mac 호환',
          desc: '삽입된 2D 구조식과 3D <code>.glb</code> 모델은 PowerPoint 표준 개체입니다. <code>func_SMILES2IMG</code>가 없는 학회장 연단 PC·macOS에서도 회전과 모핑이 동일하게 재생됩니다.'
        },
        {
          title: '⚡ 무료 MIT · 간편 설치',
          desc: 'PowerPoint, Word, Excel 통합 무료 오픈소스입니다. Windows 관리자 권한 없이 사용자 영역에 설치되고 한 줄 명령으로 제거됩니다.'
        }
      ],
      ctaTitle: 'PowerPoint · Word · Excel 통합 설치',
      ctaMeta:
        '열려 있는 오피스 창을 닫고 <strong>Windows PowerShell</strong>에 아래 명령어 한 줄을 실행한 뒤 PowerPoint를 다시 열면 <code>[SMILES]</code> 탭이 나타납니다:',
      installHref: '/ko/guide/install',
      installLabel: '상세 설치 가이드 & 사내망 오프라인 ZIP 설치 →',
      playbookHref: '/ko/powerpoint/playbook',
      playbookLabel: '8슬라이드 PowerPoint 실습 덱(Playbook .pptx) 미리보기 →'
    }
  }

  return {
    ctaClass: 'is-ppt',
    cards: [
      {
        title: '🔒 Local-First Execution',
        desc: '2D sketching (Ketcher), rendering (Indigo), SVG copy, and all 17 property lookups run locally on your machine—confidential SAR slides stay private.'
      },
      {
        title: '🎤 Safe on Any Podium PC & Mac',
        desc: '2D shapes and 3D <code>.glb</code> models are standard PowerPoint objects. Rotations and Morph transitions work across podium PCs, macOS, and iPads without installing the add-in.'
      },
      {
        title: '⚡ Free MIT · No Admin Rights Needed',
        desc: 'Free and open-source across PowerPoint, Word, and Excel. Installs to your Windows user profile without administrator privileges.'
      }
    ],
    ctaTitle: 'Install Across PowerPoint, Word & Excel',
    ctaMeta:
      'Close Office apps, paste this command into <strong>Windows PowerShell</strong> (no admin rights required), then reopen PowerPoint to unlock the <code>[SMILES]</code> tab:',
    installHref: '/guide/install',
    installLabel: 'Full Install Guide & Offline Corporate ZIP →',
    playbookHref: '/powerpoint/playbook',
    playbookLabel: 'Preview 8-Slide PowerPoint Playbook (.pptx) →'
  }
})
</script>

<template>
  <div class="s2i-overview-cta vp-raw">
    <div class="s2i-trust-strip">
      <div v-for="card in content.cards" :key="card.title" class="s2i-trust-card">
        <div class="s2i-trust-title">{{ card.title }}</div>
        <div class="s2i-trust-desc" v-html="card.desc" />
      </div>
    </div>

    <div class="s2i-install-cta" :class="content.ctaClass">
      <div class="s2i-install-cta-title">{{ content.ctaTitle }}</div>
      <p class="s2i-install-cta-meta" v-html="content.ctaMeta" />
      <QuickInstallBox inline />
      <div class="s2i-install-cta-links">
        <a :href="withBase(content.installHref)">{{ content.installLabel }}</a>
        <a :href="withBase(content.playbookHref)">{{ content.playbookLabel }}</a>
      </div>
    </div>
  </div>
</template>

<style scoped>
.s2i-trust-strip {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 12px;
  margin: 18px 0 22px;
}

.s2i-trust-card {
  padding: 13px 15px;
  border-radius: 10px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.s2i-trust-title {
  font-size: 14px;
  font-weight: 600;
  line-height: 1.4;
  color: var(--vp-c-text-1);
  margin-bottom: 4px;
}

.s2i-trust-desc {
  font-size: 13px;
  line-height: 1.5;
  color: var(--vp-c-text-2);
}

.s2i-trust-desc :deep(code),
.s2i-install-cta-meta :deep(code) {
  border-radius: 4px;
  padding: 2px 5px;
  background-color: var(--vp-code-bg);
  color: var(--vp-code-color);
  font-family: var(--vp-font-family-mono);
  font-size: var(--vp-code-font-size);
}

.s2i-install-cta {
  padding: 18px 20px;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
  margin: 20px 0 24px;
}

.s2i-install-cta.is-word {
  border-color: var(--s2i-word-border);
  background: var(--s2i-word-bg);
}

.s2i-install-cta.is-ppt {
  border-color: var(--s2i-ppt-border);
  background: var(--s2i-ppt-bg);
}

.s2i-install-cta-title {
  font-size: 16px;
  font-weight: 700;
  line-height: 1.4;
  color: var(--vp-c-text-1);
  margin: 0 0 6px;
}

.s2i-install-cta-meta {
  font-size: 13.5px;
  line-height: 1.55;
  color: var(--vp-c-text-2);
  margin: 0;
}

.s2i-install-cta-links {
  display: flex;
  flex-wrap: wrap;
  gap: 14px;
  margin-top: 12px;
  font-size: 13.5px;
  font-weight: 500;
}

.s2i-install-cta-links a {
  color: var(--vp-c-brand-1);
  text-decoration: underline;
  text-underline-offset: 2px;
  transition: color 0.2s;
}

.s2i-install-cta-links a:hover {
  color: var(--vp-c-brand-2);
}

@media (max-width: 720px) {
  .s2i-trust-strip {
    grid-template-columns: 1fr;
  }

  .s2i-install-cta {
    padding: 15px 14px;
  }
}
</style>
