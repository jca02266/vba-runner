import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve('docs');
const requiredKeys = ['type', 'title', 'description', 'tags', 'status'];
const errors = [];

function markdownFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(directory, entry.name);
    if (entry.isDirectory()) return markdownFiles(target);
    return entry.isFile() && entry.name.endsWith('.md') ? [target] : [];
  });
}

function frontmatter(content, file) {
  if (path.basename(file) === 'index.md' || path.basename(file) === 'log.md') {
    return null;
  }
  if (!content.startsWith('---\n')) {
    errors.push(`${file}: missing YAML frontmatter`);
    return null;
  }
  const end = content.indexOf('\n---\n', 4);
  if (end === -1) {
    errors.push(`${file}: unterminated YAML frontmatter`);
    return null;
  }
  const block = content.slice(4, end);
  for (const key of requiredKeys) {
    if (!new RegExp(`^${key}:`, 'm').test(block)) {
      errors.push(`${file}: missing frontmatter key '${key}'`);
    }
  }
  const staleAfter = block.match(/^stale_after:\s*(\S+)\s*$/m);
  if (staleAfter && !/^\d{4}-\d{2}-\d{2}$/.test(staleAfter[1])) {
    errors.push(`${file}: invalid stale_after '${staleAfter[1]}'`);
  }
  for (const match of block.matchAll(/^\s*resource:\s*(\S+)\s*$/gm)) {
    const resource = match[1].replace(/^['"]|['"]$/g, '');
    if (/^[a-z][a-z0-9+.-]*:/i.test(resource)) continue;
    const target = resource.startsWith('/')
      ? path.resolve(root, `.${resource}`)
      : path.resolve(path.dirname(file), resource);
    if (!fs.existsSync(target)) {
      errors.push(`${file}: source resource does not exist '${resource}'`);
    }
  }
  return block;
}

function checkLinks(content, file) {
  const directory = path.dirname(file);
  for (const match of content.matchAll(/!?\[[^\]]*\]\(([^)]+)\)/g)) {
    const raw = match[1].trim().replace(/^<|>$/g, '');
    if (!raw || raw.startsWith('#') || /^[a-z][a-z0-9+.-]*:/i.test(raw)) continue;
    const target = raw.split('#', 1)[0].split('?', 1)[0];
    if (!target) continue;
    if (!fs.existsSync(path.resolve(directory, target))) {
      errors.push(`${file}: broken link '${raw}'`);
    }
  }
}

for (const file of markdownFiles(root)) {
  const content = fs.readFileSync(file, 'utf8');
  frontmatter(content, file);
  checkLinks(content, file);
}

if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exitCode = 1;
} else {
  console.log('OKF documentation validation passed.');
}
