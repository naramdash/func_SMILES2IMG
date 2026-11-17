const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const docsDir = path.join(repoRoot, 'docs');
const publicDir = path.join(docsDir, 'public');
const scriptsDir = path.join(repoRoot, 'scripts');

const bannedScripts = [
  'Draw-ContextMenuFlyout.ps1',
  'Generate-Excel-ContextMenu-Guide.ps1',
  'Generate-Word-ContextMenu-Guide.ps1',
  'Generate-PowerPoint-ContextMenu-Guide.ps1',
  'Update-ContextMenu-Guide-English.ps1',
  'Generate-Excel-Ribbon-Guide.ps1',
  'Update-Ribbon-Guide-English.ps1',
  'Generate-Hero-Gif.ps1',
  'Generate-Ketcher-Graphic.ps1',
  'Generate-Workflow-Graphic.ps1',
  'Generate-Guide-Videos.ps1'
];

const prohibitedTerms = [
  'SMILES for Office',
  'SMILES Office Suite',
  'SMILES for Excel',
  'SMILES for Word',
  'SMILES for PowerPoint',
  'About SMILES'
];

const issues = [];

// 1. Ensure no banned synthetic/emulated scripts remain
for (const s of bannedScripts) {
  if (fs.existsSync(path.join(scriptsDir, s))) {
    issues.push(`Banned synthetic generator script still exists: scripts/${s}`);
  }
}

// 2. Validate docs/public/media/manifest.json and all feature media assets
const manifestPath = path.join(publicDir, 'media', 'manifest.json');
if (!fs.existsSync(manifestPath)) {
  issues.push('Missing live media manifest: docs/public/media/manifest.json');
} else {
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  const ids = Object.keys(manifest);
  if (ids.length < 25) {
    issues.push(`Expected >= 25 live feature media entries in manifest.json, got ${ids.length}`);
  }

  for (const id of ids) {
    const entry = manifest[id];
    const checkAsset = (relUrl, minBytes, label) => {
      if (!relUrl || typeof relUrl !== 'string') {
        issues.push(`[${id}] Missing ${label} path in manifest`);
        return;
      }
      const cleanRel = relUrl.startsWith('/') ? relUrl.substring(1) : relUrl;
      const diskPath = path.join(publicDir, cleanRel);
      if (!fs.existsSync(diskPath)) {
        issues.push(`[${id}] Missing ${label} file on disk: docs/public/${cleanRel}`);
        return;
      }
      const sz = fs.statSync(diskPath).size;
      if (sz < minBytes) {
        issues.push(`[${id}] ${label} file too small (${sz} B < ${minBytes} B): docs/public/${cleanRel}`);
      }
    };

    checkAsset(entry.shot, 5000, 'live screenshot (.png)');
    checkAsset(entry.webm, 5000, 'live video (.webm)');
    checkAsset(entry.mp4, 5000, 'live video (.mp4)');
    checkAsset(entry.vttEn, 40, 'English captions (.en.vtt)');
    checkAsset(entry.vttKo, 40, 'Korean captions (.ko.vtt)');
  }

  // 3. Verify guide markdown pages embed valid <FeatureMedia id="..." /> and zero prohibited terms
  function getMdFiles(dir) {
    let out = [];
    for (const name of fs.readdirSync(dir)) {
      if (name === '.vitepress' || name === 'node_modules') continue;
      const full = path.join(dir, name);
      if (fs.statSync(full).isDirectory()) out = out.concat(getMdFiles(full));
      else if (name.endsWith('.md')) out.push(full);
    }
    return out;
  }

  const allMd = getMdFiles(docsDir);
  const usedIds = new Set();
  for (const mdPath of allMd) {
    const rel = path.relative(docsDir, mdPath).replace(/\\/g, '/');
    const content = fs.readFileSync(mdPath, 'utf8');

    for (const term of prohibitedTerms) {
      if (content.includes(term)) {
        issues.push(`[${rel}] Contains prohibited brand/label term "${term}" (violates AGENTS.md §0)`);
      }
    }

    const fmMatches = [...content.matchAll(/<FeatureMedia\s+[^>]*id="([^"]+)"/g)];
    for (const m of fmMatches) {
      const fid = m[1];
      usedIds.add(fid);
      if (!manifest[fid]) {
        issues.push(`[${rel}] References unknown <FeatureMedia id="${fid}"> not present in manifest.json`);
      }
    }

    const isFeatureGuide =
      (rel.startsWith('guide/') || rel.startsWith('ko/guide/')) &&
      !rel.endsWith('faq.md');
    if (isFeatureGuide && fmMatches.length === 0) {
      issues.push(`[${rel}] Guide page has no <FeatureMedia> live screenshot + video components`);
    }
  }

  for (const id of ids) {
    if (!usedIds.has(id)) {
      issues.push(`[manifest.json] Feature "${id}" is captured but not embedded in any documentation page`);
    }
  }
}

if (issues.length > 0) {
  console.error(`FAIL: Found ${issues.length} docs media audit issue(s):`);
  for (const issue of issues) console.error('  - ' + issue);
  process.exit(1);
}

console.log('SUCCESS: All live screenshots, videos, captions, naming rules, and FeatureMedia embeds passed!');
