/**
 * 組み込み関数の異常値スモークファザー。
 *
 * builtinEnv に登録された全組み込み関数へ、各引数位置に敵対的な値
 * （Null / Empty / Nothing / "" / 巨大数 / 配列 など）を総当たりで注入し、
 * 「VBA エラーとして制御されず、JS の生例外（TypeError / RangeError 等）で
 * 落ちる」ケースだけを検出して報告する。
 *
 * 実行: npx tsx scripts/fuzz-builtins.ts [関数名...]
 *   引数を渡すとその関数だけをファズする（デバッグ用）。
 *
 * 判定基準:
 *   - VbaError (`e.type === 'VbaError'`) → 正常（仕様どおりの実行時エラー）
 *   - ParseError / LexError / Compile error → 正常（式として不成立なだけ）
 *   - それ以外の例外（TypeError 等の JS 生例外、engine 内部の generic Error）→ バグ候補
 *   - 正常終了でも結果文字列に "Symbol(" が漏れていたら → バグ候補（Null の生漏れ）
 */
import { evalVBASingle } from '../test-libs/test-runner';
import { ParseError } from '../src/engine/parser';
import { LexError } from '../src/engine/lexer';

// 敵対的な値（VBA 式リテラルとして注入する）
const ADVERSARIAL = [
    'Null',
    'Empty',
    'Nothing',
    '""',
    '"abc"',
    '0',
    '-1',
    '1.5',
    '2147483648',      // Long 範囲超え
    '1E308',
    'Array(1, 2, 3)',
    'True',
];

// 他の引数位置に置く無難な値
const BENIGN = '1';

// ParamArray 位置に置く追加引数の数
const PARAMARRAY_EXTRA = 1;

// 引数位置数の上限（ParamArray などで無限に増えないように）
const MAX_POSITIONS = 6;

interface Finding {
    func: string;
    expr: string;
    kind: 'js-exception' | 'symbol-leak' | 'generic-error';
    detail: string;
}

function classifyError(e: any): 'vba' | 'parse' | 'js' | 'generic' {
    if (e?.type === 'VbaError') return 'vba';
    if (e instanceof ParseError || e instanceof LexError) return 'parse';
    if (typeof e?.message === 'string' && /^(\[.*\] )?Compile error/.test(e.message)) return 'parse';
    if (e instanceof TypeError || e instanceof RangeError || e instanceof ReferenceError) return 'js';
    return 'generic';
}

function paramPositions(fn: any): { min: number; max: number } {
    const spec = fn.__vbaParamSpec__ as Array<{ optional?: boolean; isParamArray?: boolean }> | undefined;
    if (spec) {
        const min = spec.filter(p => !p.optional && !p.isParamArray).length;
        const hasParamArray = spec.some(p => p.isParamArray);
        const max = Math.min(spec.length + (hasParamArray ? PARAMARRAY_EXTRA : 0), MAX_POSITIONS);
        return { min, max };
    }
    const overloads = fn.__vbaOverloads__ as Array<{ params: any[] }> | undefined;
    if (overloads) {
        const arities = overloads.map(o => o.params.length);
        const mins = overloads.map(o => o.params.filter((p: any) => !p.optional && !p.isParamArray).length);
        return { min: Math.min(...mins), max: Math.min(Math.max(...arities), MAX_POSITIONS) };
    }
    // メタデータなし: JS 関数の宣言引数数から推定
    return { min: 0, max: Math.min(fn.length || 1, MAX_POSITIONS) };
}

function buildExpr(name: string, arity: number, injectPos: number, injectVal: string): string {
    const args: string[] = [];
    for (let i = 0; i < arity; i++) {
        args.push(i === injectPos ? injectVal : BENIGN);
    }
    return arity === 0 ? `${name}()` : `${name}(${args.join(', ')})`;
}

function main(): void {
    const only = process.argv.slice(2).map(s => s.toLowerCase());
    const findings: Finding[] = [];
    let calls = 0;
    let okCount = 0;

    // 関数名の列挙用に一時 Evaluator を作る（builtinEnv は private のためリフレクションで参照）
    const probe = evalVBASingle('', { onPrint: () => {} });
    const registry: Map<string, any> = (probe as any).builtinEnv.variables;

    const names = [...registry.keys()]
        .filter(name => typeof registry.get(name) === 'function')
        .filter(name => only.length === 0 || only.includes(name))
        .sort();

    console.log(`fuzz 対象: ${names.length} 関数 × ${ADVERSARIAL.length} 値`);

    for (const name of names) {
        const fn = registry.get(name);
        const { min, max } = paramPositions(fn);
        // 状態汚染を避けるため関数ごとに新しい Evaluator
        const ev = evalVBASingle('', { onPrint: () => {} });

        // 各引数位置に敵対値を1つずつ注入（アリティは必須数と注入位置+1 の大きい方）
        const cases: string[] = [];
        for (let pos = 0; pos < Math.max(max, 1); pos++) {
            const arity = Math.max(min, pos + 1);
            if (arity > MAX_POSITIONS) break;
            for (const v of ADVERSARIAL) {
                cases.push(buildExpr(name, arity, pos, v));
            }
        }
        // 引数ゼロ呼び出しも1回試す
        if (min === 0) cases.push(buildExpr(name, 0, -1, ''));

        for (const expr of cases) {
            calls++;
            try {
                const result = ev.evalExpression(expr);
                okCount++;
                const s = typeof result === 'string' ? result : '';
                if (s.includes('Symbol(')) {
                    findings.push({ func: name, expr, kind: 'symbol-leak', detail: `結果に生 Symbol が漏れた: ${JSON.stringify(s.slice(0, 80))}` });
                }
            } catch (e: any) {
                const cls = classifyError(e);
                if (cls === 'vba' || cls === 'parse') { okCount++; continue; }
                findings.push({
                    func: name,
                    expr,
                    kind: cls === 'js' ? 'js-exception' : 'generic-error',
                    detail: `${e?.constructor?.name ?? typeof e}: ${String(e?.message ?? e).slice(0, 160)}`,
                });
            }
        }
    }

    console.log(`\n呼び出し ${calls} 件 / 正常(VBAエラー含む) ${okCount} 件 / バグ候補 ${findings.length} 件\n`);

    if (findings.length > 0) {
        // 関数ごとにまとめて表示（同一関数の同種エラーは代表1件+件数）
        const byFunc = new Map<string, Finding[]>();
        for (const f of findings) {
            const list = byFunc.get(f.func) ?? [];
            list.push(f);
            byFunc.set(f.func, list);
        }
        for (const [func, list] of byFunc) {
            console.log(`■ ${func} (${list.length} 件)`);
            const seen = new Set<string>();
            for (const f of list) {
                const key = f.kind + '|' + f.detail;
                if (seen.has(key)) continue;
                seen.add(key);
                console.log(`   [${f.kind}] ${f.expr}`);
                console.log(`     → ${f.detail}`);
            }
            console.log();
        }
        process.exitCode = 1;
    } else {
        console.log('✅ バグ候補なし');
    }
}

main();
