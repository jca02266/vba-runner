import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

const [,, entryFile, ...args] = process.argv;

if (!entryFile) {
  console.error('Usage: node scripts/wrapper.mjs <repo-relative-entry-file> [...args]');
  process.exit(1);
}

// path.joinではなくpath.resolveを使う。
// その後でリポジトリ外への ../ による脱出を拒否する。
const cliPath = path.resolve(repoRoot, entryFile);
const repoPrefix = `${repoRoot}${path.sep}`;

if (!cliPath.startsWith(repoPrefix)) {
  console.error(`リポジトリ外のファイルは実行できません: ${entryFile}`);
  process.exit(1);
}

const invocationCwd = process.env.INIT_CWD || process.cwd();
const tsx = process.platform === 'win32' ? 'tsx.cmd' : 'tsx';

const child = spawn(tsx, [cliPath, ...args], {
  cwd: invocationCwd,
  stdio: 'inherit',
});

child.on('exit', (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
  } else {
    process.exit(code ?? 1);
  }
});
