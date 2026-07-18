/**
 * フォーマッター往復テスト。
 *
 * リポジトリ内の実 VBA コーパス（.bas / .cls）に対して:
 *   1. 原文をパース → AST-A
 *   2. format() + applyEdits() で整形
 *   3. 整形結果を再パース → AST-B（パースエラーになったら「整形がコードを壊した」）
 *   4. AST-A と AST-B を位置情報抜きで比較（不一致 = 整形が意味を変えた）
 *   5. 整形結果をもう一度整形して同一か確認（冪等性）
 *
 * 実行: npx tsx scripts/roundtrip-format.ts [ファイル...]
 *   引数なしで sample/ tests/ tools/ 配下の全 .bas/.cls を対象。不一致があれば exit 1
 */
import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';
import { format, applyEdits } from '../src/lsp/formatter';

// ---------------------------------------------------------------------------
// AST 比較（loc / line / column 等の位置情報をすべて除去して比較）
// ---------------------------------------------------------------------------
const POSITION_KEYS = new Set(['loc', 'line', 'column', 'endLine', 'endColumn', 'paramsEndColumn']);
function stripPositions(n: any): any {
    if (!n || typeof n !== 'object') return n;
    if (Array.isArray(n)) return n.map(stripPositions);
    const result: any = {};
    for (const key of Object.keys(n)) {
        if (POSITION_KEYS.has(key)) continue;
        // 識別子ケーシングの正規化はフォーマッターの意図された機能（VBA は大文字小文字無視、
        // 実 VBE も同名識別子を 1 ケーシングに強制する）のため、識別子系フィールドの
        // 文字列は小文字化して比較する
        if ((key === 'name' || key === 'objectType' || key === 'label' || key === 'className' || key === 'typeName') && typeof n[key] === 'string') {
            result[key] = n[key].toLowerCase();
            continue;
        }
        result[key] = stripPositions(n[key]);
    }
    return result;
}
const serialize = (ast: any): string => JSON.stringify(stripPositions(ast));

function parseSource(source: string, file: string): any {
    const isCls = file.endsWith('.cls');
    const opts: any = { sourceLines: source.split('\n') };
    if (isCls) {
        const lower = source.toLowerCase();
        if (!lower.includes('end class')) {
            opts.parseAsClass = path.basename(file, '.cls');
        }
    }
    return new Parser(new Lexer(source).tokenize(), opts).parse();
}

// 最初に構造が食い違う AST パスを探す（レポート用）
function findFirstDiff(a: any, b: any, p = '$'): string | null {
    if (a === b) return null;
    const ta = typeof a, tb = typeof b;
    if (ta !== 'object' || tb !== 'object' || !a || !b) {
        return serialize(a) === serialize(b) ? null : `${p}: ${JSON.stringify(a)?.slice(0, 60)} ≠ ${JSON.stringify(b)?.slice(0, 60)}`;
    }
    if (Array.isArray(a) !== Array.isArray(b)) return `${p}: 配列/非配列不一致`;
    if (Array.isArray(a)) {
        if (a.length !== b.length) return `${p}: 要素数 ${a.length} ≠ ${b.length}`;
        for (let i = 0; i < a.length; i++) {
            const d = findFirstDiff(a[i], b[i], `${p}[${i}]`);
            if (d) return d;
        }
        return null;
    }
    const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
    for (const k of keys) {
        if (POSITION_KEYS.has(k)) continue;
        const d = findFirstDiff(a[k], b[k], `${p}.${k}`);
        if (d) return d;
    }
    return null;
}

// ---------------------------------------------------------------------------
// コーパス収集
// ---------------------------------------------------------------------------
function collectFiles(dirs: string[]): string[] {
    const out: string[] = [];
    const walk = (d: string) => {
        if (!fs.existsSync(d)) return;
        for (const e of fs.readdirSync(d, { withFileTypes: true })) {
            const full = path.join(d, e.name);
            if (e.isDirectory()) {
                if (e.name === 'node_modules' || e.name.startsWith('.')) continue;
                walk(full);
            } else if (/\.(bas|cls)$/i.test(e.name)) {
                out.push(full);
            }
        }
    };
    dirs.forEach(walk);
    return out;
}

const files = process.argv.length > 2
    ? process.argv.slice(2)
    : collectFiles(['sample', 'tests', 'tools']);

// ---------------------------------------------------------------------------
// 実行
// ---------------------------------------------------------------------------
interface Finding { file: string; kind: string; detail: string }
const findings: Finding[] = [];
let checked = 0;
let skippedParse = 0;

for (const file of files) {
    const source = fs.readFileSync(file, 'utf8');

    // 原文がそもそもパースできないファイル（意図的な壊れテストデータ等）はスキップ
    let astA: any;
    try {
        astA = parseSource(source, file);
    } catch {
        skippedParse++;
        continue;
    }
    checked++;

    // 1. 整形がクラッシュしないか
    let formatted: string;
    try {
        formatted = applyEdits(source, format(source));
    } catch (e: any) {
        findings.push({ file, kind: 'format-crash', detail: String(e?.message ?? e).slice(0, 140) });
        continue;
    }

    // 2. 整形結果が再パースできるか（壊れたら整形がコードを破壊している）
    let astB: any;
    try {
        astB = parseSource(formatted, file);
    } catch (e: any) {
        findings.push({ file, kind: 'reparse-fail', detail: String(e?.message ?? e).slice(0, 140) });
        continue;
    }

    // 3. AST が意味的に同一か
    if (serialize(astA) !== serialize(astB)) {
        // findFirstDiff は正規化（位置除去・識別子小文字化）済みツリー同士で比較する
        const diff = findFirstDiff(stripPositions(astA), stripPositions(astB)) ?? '(差分位置特定失敗)';
        findings.push({ file, kind: 'ast-changed', detail: diff.slice(0, 200) });
        continue;
    }

    // 4. 冪等性: 2 回目の整形が no-op か
    try {
        const formatted2 = applyEdits(formatted, format(formatted));
        if (formatted2 !== formatted) {
            // 最初に食い違う行を特定
            const l1 = formatted.split('\n'), l2 = formatted2.split('\n');
            let i = 0;
            while (i < Math.max(l1.length, l2.length) && l1[i] === l2[i]) i++;
            findings.push({
                file, kind: 'not-idempotent',
                detail: `line ${i + 1}: ${JSON.stringify(l1[i] ?? '(なし)')?.slice(0, 80)} → ${JSON.stringify(l2[i] ?? '(なし)')?.slice(0, 80)}`,
            });
        }
    } catch (e: any) {
        findings.push({ file, kind: 'format2-crash', detail: String(e?.message ?? e).slice(0, 140) });
    }
}

console.log(`検査 ${checked} ファイル（パース不能スキップ ${skippedParse}）/ 問題 ${findings.length} 件\n`);
if (findings.length > 0) {
    const byKind = new Map<string, Finding[]>();
    for (const f of findings) {
        const list = byKind.get(f.kind) ?? [];
        list.push(f);
        byKind.set(f.kind, list);
    }
    for (const [kind, list] of byKind) {
        console.log(`■ ${kind} (${list.length} 件)`);
        for (const f of list.slice(0, 10)) {
            console.log(`   ${f.file}`);
            console.log(`     → ${f.detail}`);
        }
        if (list.length > 10) console.log(`   ... 他 ${list.length - 10} 件`);
        console.log();
    }
    process.exitCode = 1;
} else {
    console.log('✅ 全ファイルで往復一致・冪等');
}
