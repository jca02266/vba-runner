/**
 * メタモルフィックテスト — オラクル不要の恒等式でサイレントバグ（値の誤り）を検出する。
 *
 * scripts/fuzz-builtins.ts が「JS 生例外で落ちる」バグ専用なのに対し、こちらは
 * 「落ちないが値が間違っている」バグを狙う。VBA 仕様上必ず成り立つ恒等式
 * （例: Mid(s,1,Len(s)) = s、(a\b)*b + (a Mod b) = a、DateAdd の逆演算）に
 * ランダム入力を流し、破れたら報告する。
 *
 * 実行: npx tsx scripts/metamorphic-test.ts [seed] [iterations]
 *   - seed 省略時は現在時刻。違反再現のため seed は常に出力される
 *   - iterations は各恒等式あたりの試行回数（デフォルト 200）
 *   - 違反が 1 件でもあれば exit 1
 */
import { evalVBASingle } from '../test-libs/test-runner';
import { VbaBoolean, VbaDate } from '../src/engine/vba-types';

// ---------------------------------------------------------------------------
// 乱数（再現可能な LCG）
// ---------------------------------------------------------------------------
let seed = process.argv[2] ? Number(process.argv[2]) : Date.now() % 2147483647;
const ITER = process.argv[3] ? Number(process.argv[3]) : 200;
let state = seed;
const rnd = (): number => {
    state = (state * 48271) % 2147483647;
    return state / 2147483647;
};
const randInt = (min: number, max: number): number => Math.floor(rnd() * (max - min + 1)) + min;
const pick = <T>(arr: T[]): T => arr[randInt(0, arr.length - 1)];

// ---------------------------------------------------------------------------
// VBA リテラル生成
// ---------------------------------------------------------------------------
const CHARSET = [
    ...'abcxyzABCXYZ019', ' ', '"', "'", '#', '&', '(', ')', '*', ',', '-', '.', '/', ':', ';',
    '<', '=', '>', '?', '\t', 'あ', '漢', '　', 'ｱ', 'Ａ',
];
const randStr = (maxLen = 12): string => {
    const len = randInt(0, maxLen);
    let s = '';
    for (let i = 0; i < len; i++) s += pick(CHARSET);
    return s;
};
const strLit = (s: string): string => '"' + s.replace(/"/g, '""') + '"';
// 指数表記にならない安全な小数リテラル
const randDbl = (): number => {
    const specials = [0, 0.5, -0.5, 1.5, 2.5, -2.5, 0.0001, 123456.789];
    if (rnd() < 0.2) return pick(specials);
    return randInt(-1000000, 1000000) + randInt(0, 9999) / 10000;
};
const numLit = (n: number): string => {
    const s = n.toFixed(10).replace(/\.?0+$/, '');
    return s === '' || s === '-' ? '0' : s;
};

// ---------------------------------------------------------------------------
// 値の比較
// ---------------------------------------------------------------------------
const norm = (v: any): any => {
    if (v instanceof VbaBoolean) return v.value !== 0;
    if (v instanceof VbaDate) return { __date__: v.value };
    if (typeof v === 'bigint') return Number(v);
    return v;
};
const eq = (a: any, b: any, relTol = 0): boolean => {
    const x = norm(a), y = norm(b);
    if (typeof x === 'object' && x?.__date__ !== undefined) {
        return typeof y === 'object' && y?.__date__ !== undefined && Math.abs(x.__date__ - y.__date__) < 1e-9;
    }
    if (typeof x === 'number' && typeof y === 'number') {
        if (x === y) return true;
        if (relTol > 0) return Math.abs(x - y) <= relTol * Math.max(1, Math.abs(x), Math.abs(y));
        return false;
    }
    return x === y;
};
const show = (v: any): string => {
    const n = norm(v);
    if (typeof n === 'object' && n?.__date__ !== undefined) return `Date(${n.__date__})`;
    return JSON.stringify(n);
};

// ---------------------------------------------------------------------------
// 恒等式の定義
// ---------------------------------------------------------------------------
interface Violation { relation: string; detail: string }
const violations: Violation[] = [];
const ev = evalVBASingle('', { onPrint: () => {} });

/** expr1 と expr2 を評価して一致を確認。恒等式が成り立たなければ違反として記録 */
function check(relation: string, expr1: string, expr2: string, relTol = 0): void {
    let v1: any, v2: any;
    try {
        v1 = ev.evalExpression(expr1);
        v2 = ev.evalExpression(expr2);
    } catch (e: any) {
        violations.push({ relation, detail: `評価エラー: ${expr1}  /  ${expr2} → ${String(e?.message ?? e).slice(0, 120)}` });
        return;
    }
    if (!eq(v1, v2, relTol)) {
        violations.push({ relation, detail: `${expr1} = ${show(v1)}  ≠  ${expr2} = ${show(v2)}` });
    }
}

/** 期待値を JS 側で計算して比較 */
function checkVal(relation: string, expr: string, expected: any, relTol = 0): void {
    let v: any;
    try {
        v = ev.evalExpression(expr);
    } catch (e: any) {
        violations.push({ relation, detail: `評価エラー: ${expr} → ${String(e?.message ?? e).slice(0, 120)}` });
        return;
    }
    if (!eq(v, expected, relTol)) {
        violations.push({ relation, detail: `${expr} = ${show(v)}  ≠  期待値 ${show(expected)}` });
    }
}

const relations: Array<{ name: string; run: () => void }> = [
    // ---- 文字列 ----
    {
        name: 'Mid(s,1,Len(s)) = s',
        run: () => { const s = strLit(randStr()); check('Mid 全体切り出し', `Mid(${s}, 1, Len(${s}))`, s); },
    },
    {
        name: 'Left(s,k) & Mid(s,k+1) = s',
        run: () => {
            const raw = randStr(); const s = strLit(raw); const k = randInt(0, raw.length);
            check('Left/Mid 分割再結合', `Left(${s}, ${k}) & Mid(${s}, ${k + 1})`, s);
        },
    },
    {
        name: 'Len(s & t) = Len(s) + Len(t)',
        run: () => {
            const a = randStr(), b = randStr();
            checkVal('Len 加法性', `Len(${strLit(a)} & ${strLit(b)})`, a.length + b.length);
        },
    },
    {
        name: 'StrReverse(StrReverse(s)) = s',
        run: () => { const s = strLit(randStr()); check('StrReverse 対合', `StrReverse(StrReverse(${s}))`, s); },
    },
    {
        name: 'Trim(s) = LTrim(RTrim(s))',
        run: () => { const s = strLit(randStr()); check('Trim 分解', `Trim(${s})`, `LTrim(RTrim(${s}))`); },
    },
    {
        name: 'Join(Split(s,d),d) = s',
        run: () => {
            const s = strLit(randStr());
            const d = strLit(pick([',', ';', '--', 'あ', ' ']));
            check('Split/Join 往復', `Join(Split(${s}, ${d}), ${d})`, s);
        },
    },
    {
        name: 'Space(n) = String(n," ")',
        run: () => { const n = randInt(0, 40); check('Space/String', `Space(${n})`, `String(${n}, " ")`); },
    },
    {
        name: 'Asc(Chr(k)) = k',
        run: () => { const k = randInt(1, 255); checkVal('Asc/Chr 往復', `Asc(Chr(${k}))`, k); },
    },
    {
        name: 'Replace(s,t,t) = s',
        run: () => {
            const s = strLit(randStr());
            const t = strLit(randStr(3));
            check('Replace 恒等置換', `Replace(${s}, ${t}, ${t})`, s);
        },
    },
    {
        name: 'InStr ヒット位置の Mid は一致',
        run: () => {
            const raw = randStr(8) + pick(['ab', 'xy', '漢あ']) + randStr(4);
            const t = pick(['ab', 'xy', '漢あ']);
            const s = strLit(raw), tl = strLit(t);
            const posExpr = `InStr(${s}, ${tl})`;
            let pos: any;
            try { pos = ev.evalExpression(posExpr); } catch { return; }
            if (typeof pos === 'number' && pos > 0) {
                check('InStr/Mid 整合', `Mid(${s}, ${posExpr}, Len(${tl}))`, tl);
            }
        },
    },
    {
        name: 'UCase 冪等',
        run: () => { const s = strLit(randStr()); check('UCase 冪等', `UCase(UCase(${s}))`, `UCase(${s})`); },
    },
    {
        name: 'StrComp 反対称',
        run: () => {
            const a = strLit(randStr(6)), b = strLit(randStr(6));
            checkVal('StrComp 反対称', `StrComp(${a}, ${b}, 0) + StrComp(${b}, ${a}, 0)`, 0);
        },
    },
    // ---- 数値 ----
    {
        name: '(a\\b)*b + (a Mod b) = a',
        run: () => {
            const a = randInt(-100000, 100000);
            let b = randInt(-1000, 1000); if (b === 0) b = 7;
            checkVal('整数除算と Mod の整合', `(${a} \\ ${b}) * ${b} + (${a} Mod ${b})`, a);
        },
    },
    {
        name: 'Abs(-x) = Abs(x)',
        run: () => { const x = numLit(randDbl()); check('Abs 偶関数', `Abs(-(${x}))`, `Abs(${x})`); },
    },
    {
        name: 'Sgn(x)*Abs(x) = x',
        run: () => { const x = randDbl(); checkVal('Sgn*Abs 復元', `Sgn(${numLit(x)}) * Abs(${numLit(x)})`, x); },
    },
    {
        name: 'Fix(x) = Sgn(x)*Int(Abs(x))',
        run: () => { const x = numLit(randDbl()); check('Fix/Int 関係', `Fix(${x})`, `Sgn(${x}) * Int(Abs(${x}))`); },
    },
    {
        name: 'Int(x) <= x < Int(x)+1',
        run: () => {
            const x = randDbl();
            let intX: any;
            try { intX = Number(ev.evalExpression(`Int(${numLit(x)})`)); } catch { return; }
            if (!(intX <= x && x < intX + 1)) {
                violations.push({ relation: 'Int 床関数性', detail: `Int(${numLit(x)}) = ${intX} が [x-1 < Int(x) <= x] を満たさない` });
            }
        },
    },
    {
        name: 'Val("&H" & Hex(a)) = a',
        run: () => { const a = randInt(0, 2147483647); checkVal('Hex/Val 往復', `Val("&H" & Hex(${a}))`, a); },
    },
    {
        name: 'Val(Str(a)) = a',
        run: () => { const a = randInt(-2147483647, 2147483647); checkVal('Str/Val 往復', `Val(Str(${a}))`, a); },
    },
    {
        name: 'CLng(CStr(a)) = a',
        run: () => { const a = randInt(-2147483647, 2147483647); checkVal('CStr/CLng 往復', `CLng(CStr(${a}))`, a); },
    },
    {
        name: 'CDbl(CStr(x)) ≈ x',
        run: () => { const x = randDbl(); checkVal('CStr/CDbl 往復', `CDbl(CStr(${numLit(x)}))`, x, 1e-13); },
    },
    {
        name: 'Sqr(x)^2 ≈ x',
        run: () => { const x = Math.abs(randDbl()); checkVal('Sqr/^2 往復', `Sqr(${numLit(x)}) ^ 2`, x, 1e-12); },
    },
    {
        name: '|CInt(x) - x| <= 0.5',
        run: () => {
            const x = randInt(-30000, 30000) + randInt(0, 99) / 100;
            let ci: any;
            try { ci = Number(ev.evalExpression(`CInt(${numLit(x)})`)); } catch { return; }
            if (Math.abs(ci - x) > 0.5 + 1e-9) {
                violations.push({ relation: 'CInt 距離', detail: `CInt(${numLit(x)}) = ${ci}（|差| > 0.5）` });
            }
        },
    },
    // ---- 日付 ----
    {
        name: 'DateAdd("d",-n,DateAdd("d",n,dt)) = dt',
        run: () => {
            const y = randInt(1950, 2100), m = randInt(1, 12), d = randInt(1, 28), n = randInt(0, 5000);
            const dt = `DateSerial(${y}, ${m}, ${d})`;
            check('DateAdd 逆演算', `DateAdd("d", -${n}, DateAdd("d", ${n}, ${dt}))`, dt);
        },
    },
    {
        name: 'DateDiff("d", dt, DateAdd("d",n,dt)) = n',
        run: () => {
            const y = randInt(1950, 2100), m = randInt(1, 12), d = randInt(1, 28), n = randInt(0, 5000);
            const dt = `DateSerial(${y}, ${m}, ${d})`;
            checkVal('DateDiff/DateAdd 整合', `DateDiff("d", ${dt}, DateAdd("d", ${n}, ${dt}))`, n);
        },
    },
    {
        name: 'DateSerial 成分往復',
        run: () => {
            const y = randInt(100, 9999), m = randInt(1, 12), d = randInt(1, 28);
            const dt = `DateSerial(${y}, ${m}, ${d})`;
            checkVal('Year 往復', `Year(${dt})`, y);
            checkVal('Month 往復', `Month(${dt})`, m);
            checkVal('Day 往復', `Day(${dt})`, d);
        },
    },
    {
        name: 'TimeSerial 成分往復',
        run: () => {
            const h = randInt(0, 23), n = randInt(0, 59), s = randInt(0, 59);
            const t = `TimeSerial(${h}, ${n}, ${s})`;
            checkVal('Hour 往復', `Hour(${t})`, h);
            checkVal('Minute 往復', `Minute(${t})`, n);
            checkVal('Second 往復', `Second(${t})`, s);
        },
    },
    {
        name: 'Weekday は 7 日周期',
        run: () => {
            const y = randInt(1950, 2100), m = randInt(1, 12), d = randInt(1, 28);
            const fdow = randInt(1, 7);
            const dt = `DateSerial(${y}, ${m}, ${d})`;
            check('Weekday 周期性', `Weekday(DateAdd("d", 7, ${dt}), ${fdow})`, `Weekday(${dt}, ${fdow})`);
        },
    },
    {
        name: 'IsDate(CStr(date)) = True',
        run: () => {
            const y = randInt(1950, 2100), m = randInt(1, 12), d = randInt(1, 28);
            checkVal('CStr した日付は IsDate', `IsDate(CStr(DateSerial(${y}, ${m}, ${d})))`, true);
        },
    },
    // ---- Currency / Decimal（内部 BigInt 経路）----
    {
        name: 'CCur の正確な加法',
        run: () => {
            // 4 桁固定小数の加算は Currency では正確（double の丸め誤差が出たらバグ）
            const a = randInt(-99999999, 99999999), b = randInt(-99999999, 99999999);
            const as = `${Math.trunc(a / 10000)}.${String(Math.abs(a % 10000)).padStart(4, '0')}`;
            const bs = `${Math.trunc(b / 10000)}.${String(Math.abs(b % 10000)).padStart(4, '0')}`;
            const sum = a + b;
            const ss = `${Math.trunc(sum / 10000)}.${String(Math.abs(sum % 10000)).padStart(4, '0')}`;
            check('CCur 加法の正確性', `CStr(CCur("${as}") + CCur("${bs}"))`, `CStr(CCur("${ss}"))`);
        },
    },
    {
        name: 'CCur(CStr(cur)) 往復',
        run: () => {
            const a = randInt(-999999999, 999999999);
            const as = `${Math.trunc(a / 10000)}.${String(Math.abs(a % 10000)).padStart(4, '0')}`;
            check('CCur/CStr 往復', `CStr(CCur(CStr(CCur("${as}"))))`, `CStr(CCur("${as}"))`);
        },
    },
    {
        name: 'CDec の正確な加法',
        run: () => {
            const a = randInt(-999999, 999999), b = randInt(-999999, 999999);
            const as = `${Math.trunc(a / 100)}.${String(Math.abs(a % 100)).padStart(2, '0')}`;
            const bs = `${Math.trunc(b / 100)}.${String(Math.abs(b % 100)).padStart(2, '0')}`;
            const sum = a + b;
            const ss = `${Math.trunc(sum / 100)}.${String(Math.abs(sum % 100)).padStart(2, '0')}`;
            check('CDec 加法の正確性', `CStr(CDec("${as}") + CDec("${bs}"))`, `CStr(CDec("${ss}"))`);
        },
    },
    // ---- 変換・その他 ----
    {
        name: 'Val("&O" & Oct(a)) = a',
        run: () => { const a = randInt(0, 2147483647); checkVal('Oct/Val 往復', `Val("&O" & Oct(${a}))`, a); },
    },
    {
        name: '文字列同士の + と & は等価',
        run: () => {
            const a = strLit(randStr()), b = strLit(randStr());
            check('文字列 + と &', `${a} + ${b}`, `${a} & ${b}`);
        },
    },
    {
        name: 'CBool(CStr(b)) = b',
        run: () => {
            const b = pick(['True', 'False']);
            check('CBool/CStr 往復', `CBool(CStr(${b}))`, b);
        },
    },
    {
        name: 'IsNumeric(CStr(n)) = True',
        run: () => { const x = randDbl(); checkVal('CStr した数値は IsNumeric', `IsNumeric(CStr(${numLit(x)}))`, true); },
    },
    {
        name: 'Log(Exp(x)) ≈ x',
        run: () => {
            const x = randInt(-50, 50) + randInt(0, 999) / 1000;
            checkVal('Log/Exp 往復', `Log(Exp(${numLit(x)}))`, x, 1e-12);
        },
    },
    {
        name: 'Sin^2 + Cos^2 ≈ 1',
        run: () => {
            const x = numLit(randDbl());
            checkVal('三角関数の恒等式', `Sin(${x}) ^ 2 + Cos(${x}) ^ 2`, 1, 1e-12);
        },
    },
    {
        name: 'Hex(Val("&H" & h)) 往復',
        run: () => {
            const a = randInt(1, 2147483647);
            check('Val/Hex 往復', `Hex(Val("&H" & Hex(${a})))`, `Hex(${a})`);
        },
    },
    {
        name: 'Filter の結果はすべてマッチを含む',
        run: () => {
            const words = Array.from({ length: randInt(1, 6) }, () => randStr(6));
            const t = pick(['a', 'x', '1', 'あ']);
            const arrExpr = `Array(${words.map(w => strLit(w)).join(', ') || '""'})`;
            const expected = words.filter(w => w.includes(t)).length;
            // Filter は一致 0 件時に UBound = -1 の空配列を返す
            checkVal('Filter 件数整合', `UBound(Filter(${arrExpr}, ${strLit(t)})) + 1`, expected);
        },
    },
];

// ---------------------------------------------------------------------------
// 実行
// ---------------------------------------------------------------------------
// --selftest: 検出器の健全性確認。既知の偽恒等式が違反として報告されなければ異常
if (process.argv.includes('--selftest')) {
    check('SELFTEST(必ず違反になるべき)', '1 + 1', '3');
    check('SELFTEST(文字列)', '"abc"', '"abd"');
    if (violations.length === 2) {
        console.log('✅ selftest OK: 偽恒等式 2 件が正しく違反として検出された');
        process.exit(0);
    }
    console.error(`❌ selftest FAILED: 違反 ${violations.length} 件（期待 2 件）— 検出器が壊れている`);
    process.exit(1);
}

console.log(`メタモルフィックテスト: ${relations.length} 恒等式 × ${ITER} 試行 (seed=${seed})`);

for (const rel of relations) {
    const before = violations.length;
    for (let i = 0; i < ITER; i++) {
        rel.run();
        // 同じ恒等式で違反が溢れたら打ち切り（レポートを読みやすく保つ）
        if (violations.length - before >= 5) break;
    }
}

const total = relations.length * ITER;
console.log(`\n試行 約${total} 件 / 違反 ${violations.length} 件\n`);

if (violations.length > 0) {
    const byRel = new Map<string, string[]>();
    for (const v of violations) {
        const list = byRel.get(v.relation) ?? [];
        if (list.length < 5) list.push(v.detail);
        byRel.set(v.relation, list);
    }
    for (const [rel, details] of byRel) {
        console.log(`■ ${rel}`);
        for (const d of details) console.log(`   ${d}`);
        console.log();
    }
    console.log(`再現: npx tsx scripts/metamorphic-test.ts ${seed} ${ITER}`);
    process.exitCode = 1;
} else {
    console.log('✅ 違反なし');
}
