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
    for (const rawLine of fs.readFileSync('scripts/diff-allowlist.txt', 'utf8').split('\n')) {
        // 行頭コメント行はスキップ。行内の "  # ..." 以降は行末コメントとして切り捨てる
        const line = rawLine.trim();
        if (!line || line.startsWith('#')) continue;
        const id = line.split(/\s+#/)[0].trim();
        if (id) allowlist.add(id);
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

// 不一致をパターン分類して要約（--full で全件一覧も出す）
if (mismatches.length > 0) {
    const categorize = (id: string): string => {
        const ra = a.get(id), rb = b.get(id);
        if (!ra || !rb) return 'MISSING(片側にのみ存在)';
        if (ra.status !== rb.status) return `STATUS(${ra.status} vs ${rb.status})`;
        if (ra.status === 'ERR') return `ERRNUM(${ra.payload} vs ${rb.payload})`;
        const [ta, ...vaRest] = ra.payload.split(':');
        const [tb, ...vbRest] = rb.payload.split(':');
        const va = vaRest.join(':'), vb = vbRest.join(':');
        if (ta !== tb && va === vb) return `TYPETAG-ONLY(${ta} vs ${tb})`;
        if (ta !== tb) return `TYPE+VALUE(${ta} vs ${tb})`;
        return `VALUE(${ta})`;
    };
    const byCat = new Map<string, string[]>();
    for (const m of mismatches) {
        const id = m.split(':')[0];
        const cat = categorize(id);
        (byCat.get(cat) ?? byCat.set(cat, []).get(cat)!).push(m);
    }
    console.log('■ 不一致のパターン分類（要仕様裁定）');
    for (const [cat, list] of [...byCat.entries()].sort((x, y) => y[1].length - x[1].length)) {
        console.log(`\n  [${list.length} 件] ${cat}`);
        for (const m of list.slice(0, 5)) console.log(`    ${m}`);
        if (list.length > 5) console.log(`    ... 他 ${list.length - 5} 件`);
    }
    if (process.argv.includes('--full')) {
        console.log('\n■ 全件一覧');
        for (const m of mismatches) console.log(`  ${m}`);
    } else {
        console.log('\n（全件一覧は --full を付けて実行）');
    }
    process.exitCode = 1;
} else {
    console.log('✅ 許容外の不一致なし');
}
if (allowed.length > 0) {
    console.log(`\n（許容リスト適用済み ${allowed.length} 件は scripts/diff-allowlist.txt を参照）`);
}
