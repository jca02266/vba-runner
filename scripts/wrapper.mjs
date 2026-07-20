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

// 実行対象をリポジトリルートからの相対パスとして解決する。
const cliPath = path.resolve(repoRoot, entryFile);
const repoPrefix = `${repoRoot}${path.sep}`;

if (!cliPath.startsWith(repoPrefix)) {
  console.error(`リポジトリ外のファイルは実行できません: ${entryFile}`);
  process.exit(1);
}

// npmはスクリプトをpackage.jsonのあるディレクトリで実行するため、process.cwd()はリポジトリルートになる。
// INIT_CWDは、ユーザーがnpmコマンドを実行した元のディレクトリをnpmが自動設定する環境変数。
// 相対パスをユーザーの作業ディレクトリ基準で解決するため、CLIプロセスのcwdとしてINIT_CWDを使用する。
const invocationCwd = process.env.INIT_CWD || process.cwd();

// Node自身からtsxをロードすることで、tsx / tsx.cmdのOS差異をなくす。
const child = spawn(
  process.execPath,
  ['--import', 'tsx', cliPath, ...args],
  {
    cwd: invocationCwd,
    stdio: 'inherit',
  },
);

child.on('error', (error) => {
  console.error(`CLIの起動に失敗しました: ${error.message}`);
  process.exit(1);
});

child.on('exit', (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
  } else {
    process.exit(code ?? 1);
  }
});
