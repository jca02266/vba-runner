/**
 * 実 VBA 差分テストの vba-runner 側実行。
 *
 * scripts/gen-diff-corpus.ts が生成した diff-vba/ の全モジュールを vba-runner で
 * ロードして DiffMain.RunAll を実行し、結果を diff-vba/results-runner.txt に書き出す。
 *
 * 実行: npx tsx scripts/run-diff-local.ts
 *
 * 実 Excel 側の結果（vba-diff-results.txt）が手に入ったら:
 *   npx tsx scripts/diff-results.ts diff-vba/results-runner.txt <excel側の結果ファイル>
 */
import * as fs from 'fs';
import * as path from 'path';
import { evalVBAModules } from '../test-libs/test-runner';

const DIR = 'diff-vba';
if (!fs.existsSync(path.join(DIR, 'DiffMain.bas'))) {
    console.error('diff-vba/ がありません。先に npx tsx scripts/gen-diff-corpus.ts を実行してください');
    process.exit(1);
}

const modules = fs.readdirSync(DIR)
    .filter(f => f.endsWith('.bas'))
    .map(f => ({
        name: path.basename(f, '.bas'),
        code: fs.readFileSync(path.join(DIR, f), 'utf8'),
    }));

console.log(`${modules.length} モジュールをロードして実行中...`);
const ev = evalVBAModules(modules, { onPrint: () => {} });
ev.callProcedure('RunAll', []);

const out = (ev as any).fs.readFileSync('/sandbox/vba-diff-results.txt').toString();
const outPath = path.join(DIR, 'results-runner.txt');
fs.writeFileSync(outPath, out);

const lines = out.split(/\r?\n/).filter((l: string) => l.trim() !== '');
console.log(`完了: ${lines.length} 件の結果 → ${outPath}`);
