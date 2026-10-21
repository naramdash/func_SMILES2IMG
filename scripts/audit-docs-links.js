const fs = require('fs');
const path = require('path');

const docsDir = path.resolve(__dirname, '../docs');
const publicDir = path.resolve(docsDir, 'public');

function getAllMdFiles(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  for (const file of list) {
    if (file === '.vitepress' || file === 'node_modules') continue;
    const fullPath = path.join(dir, file);
    const stat = fs.statSync(fullPath);
    if (stat && stat.isDirectory()) {
      results = results.concat(getAllMdFiles(fullPath));
    } else if (file.endsWith('.md')) {
      results.push(fullPath);
    }
  }
  return results;
}

const mdFiles = getAllMdFiles(docsDir);
console.log(`Auditing ${mdFiles.length} markdown documentation files...`);

let issues = [];

for (const filePath of mdFiles) {
  const relPath = path.relative(docsDir, filePath);
  let content = fs.readFileSync(filePath, 'utf8');

  // Strip code blocks so SMARTS like `[CX3H1](=O)` aren't mistaken for links
  content = content.replace(/```[\s\S]*?```/g, '');
  content = content.replace(/`[^`\n]+`/g, '');

  // Regex 1: Markdown links [text](url)
  const mdRegex = /\[([^\]]*)\]\(([^)#?]+)(?:[#?][^)]*)?\)/g;
  let match;
  while ((match = mdRegex.exec(content)) !== null) {
    checkLink(relPath, filePath, match[2].trim(), 'Markdown');
  }

  // Regex 2: HTML href="url"
  const hrefRegex = /href="([^"#?]+)(?:[#?][^"]*)?"/g;
  while ((match = hrefRegex.exec(content)) !== null) {
    checkLink(relPath, filePath, match[1].trim(), 'HTML');
  }
}

function checkLink(relDoc, fullDocPath, url, type) {
  if (url.startsWith('http://') || url.startsWith('https://') || url.startsWith('mailto:') || url.startsWith('#')) {
    return;
  }

  // If base url: /func_SMILES2IMG/...
  let cleanUrl = url;
  if (cleanUrl.startsWith('/func_SMILES2IMG/')) {
    cleanUrl = cleanUrl.substring('/func_SMILES2IMG/'.length);
    const inPublic = path.join(publicDir, cleanUrl);
    const inDocs = path.join(docsDir, cleanUrl);
    const inDocsMd = path.join(docsDir, cleanUrl + '.md');
    const inDocsIndex = path.join(docsDir, cleanUrl, 'index.md');

    if (!fs.existsSync(inPublic) && !fs.existsSync(inDocs) && !fs.existsSync(inDocsMd) && !fs.existsSync(inDocsIndex)) {
      issues.push(`[${relDoc}] ${type} broken link: ${url}`);
    }
    return;
  }

  if (cleanUrl.startsWith('/')) {
    cleanUrl = cleanUrl.substring(1);
    const inPublic = path.join(publicDir, cleanUrl);
    const inDocs = path.join(docsDir, cleanUrl);
    const inDocsMd = path.join(docsDir, cleanUrl + '.md');
    const inDocsIndex = path.join(docsDir, cleanUrl, 'index.md');

    if (!fs.existsSync(inPublic) && !fs.existsSync(inDocs) && !fs.existsSync(inDocsMd) && !fs.existsSync(inDocsIndex)) {
      issues.push(`[${relDoc}] ${type} broken root link: ${url}`);
    }
    return;
  }

  // Relative link
  const docDir = path.dirname(fullDocPath);
  const relTarget = path.resolve(docDir, cleanUrl);
  const relTargetMd = path.resolve(docDir, cleanUrl + '.md');
  const relTargetIndex = path.resolve(docDir, cleanUrl, 'index.md');

  if (!fs.existsSync(relTarget) && !fs.existsSync(relTargetMd) && !fs.existsSync(relTargetIndex)) {
    issues.push(`[${relDoc}] ${type} broken relative link: ${url}`);
  }
}

if (issues.length === 0) {
  console.log('✅ SUCCESS: All internal links and file downloads resolved properly!');
} else {
  console.log(`❌ Found ${issues.length} broken links:`);
  issues.forEach(i => console.log('  ' + i));
  process.exit(1);
}
