/**
 * 実 VBA 差分テストの結果比較。
 *
 * 実行: npx tsx scripts/diff-results.ts <runner側結果> <実Excel側結果>
 *   例: npx tsx scripts/diff-results.ts diff-vba/results-runner.txt vba-diff-results.txt
 *
 * - 各行は「id \t OK|ERR \t 正規化値またはエラー番号」
 * - diff-vba/manifest.json から式を引いてレポートに添える
 * - scripts/diff-allowlist.txt（1 行 1 id、# コメント可）に載っている不一致は
 *   「許容済み」として集計から除外する（サンドボックス・スタブ由来の意図的差分用）
 * - 許容外の不一致が 1 件でもあれば exit 1
 */
import * as fs from 'fs';

const [fileA, fileB] = process.argv.slice(2);
if (!fileA || !fileB) {
    console.error('使い方: npx tsx scripts/diff-results.ts <runner側結果> <実Excel側結果>');
    process.exit(2);
}

function parse(file: string): Map<string, { status: string; payload: string }> {
    const map = new Map<string, { status: string; payload: string }>();
    for (const line of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
        if (line.trim() === '') continue;
        const [id, status, ...rest] = line.split('\t');
        map.set(id, { status, payload: rest.join('\t') });
    }
    return map;
}

const a = parse(fileA);
const b = parse(fileB);
const manifest = fs.existsSync('diff-vba/manifest.json')
    ? JSON.parse(fs.readFileSync('diff-vba/manifest.json', 'utf8'))
    : { cases: {} };

const allowlist = new Set<string>();
if (fs.existsSync('scripts/diff-allowlist.txt')) {
    for (const line of fs.readFileSync('scripts/diff-allowlist.txt', 'utf8').split('\n')) {
        const t = line.trim();
        if (t && !t.startsWith('#')) allowlist.add(t);
    }
}

const ids = [...new Set([...a.keys(), ...b.keys()])].sort();
let match = 0;
const mismatches: string[] = [];
const allowed: string[] = [];

for (const id of ids) {
    const ra = a.get(id);
    const rb = b.get(id);
    const expr = manifest.cases[id] ?? '(式不明)';
    if (!ra || !rb) {
        const msg = `${id}: 片側にのみ存在 (runner=${ra ? '有' : '無'} / excel=${rb ? '有' : '無'})  式: ${expr}`;
        (allowlist.has(id) ? allowed : mismatches).push(msg);
        continue;
    }
    if (ra.status === rb.status && ra.payload === rb.payload) {
        match++;
        continue;
    }
    const msg = `${id}: runner=[${ra.status} ${ra.payload}] excel=[${rb.status} ${rb.payload}]  式: ${expr}`;
    (allowlist.has(id) ? allowed : mismatches).push(msg);
}

console.log(`比較 ${ids.length} 件: 一致 ${match} / 不一致 ${mismatches.length} / 許容済み ${allowed.length}\n`);
if (mismatches.length > 0) {
    console.log('■ 不一致（要仕様裁定）');
    for (const m of mismatches) console.log(`  ${m}`);
    process.exitCode = 1;
} else {
    console.log('✅ 許容外の不一致なし');
}
if (allowed.length > 0) {
    console.log(`\n（許容リスト適用済み ${allowed.length} 件は scripts/diff-allowlist.txt を参照）`);
}
