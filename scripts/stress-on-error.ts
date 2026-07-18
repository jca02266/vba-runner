/**
 * On Error 経路の組み合わせストレステスト。
 *
 * ハンドラーモード × エラー発生位置 × エラー種別 × 構造（インライン / For ループ内 /
 * 子プロシージャ経由）の全組み合わせで小さな VBA プログラムを生成・実行し、
 * Debug.Print の実行トレースを MS-VBAL §5.4.4 のセマンティクスに基づく期待値モデルと
 * 突き合わせる。既存の tests/spec/ の On Error テスト群（単機能検証）が届かない
 * 「機能の相互作用」を狙う。
 *
 * 実行: npx tsx scripts/stress-on-error.ts
 *   - 全組み合わせ決定的に実行（乱数なし）。不一致が 1 件でもあれば exit 1
 *
 * モデルの前提（MS-VBAL §5.4.4 / VBA ドキュメント）:
 * - On Error Resume Next: エラー文を中断し次の文へ。Err は保持される
 * - On Error GoTo label: ハンドラーへジャンプ。Resume Next で失敗文の次へ戻り Err クリア
 * - ハンドラー実行中の再エラーは呼び出し元へ伝播（fatal）
 * - Resume / Exit Sub・Exit Function・Exit Property / 任意の On Error 文で Err は
 *   自動クリアされる（§6.1.3.2 Clear）。**End Sub はこのリストに含まれない**ため、
 *   End Sub で終わる子プロシージャが処理したエラーの Err は呼び出し元にも見える
 * - ハンドラー本体が End Sub に到達したら正常終了
 */
import { evalVBASingle } from '../test-libs/test-runner';

// ---------------------------------------------------------------------------
// シナリオ定義
// ---------------------------------------------------------------------------
type Mode = 'none' | 'resumenext' | 'goto-resumenext' | 'goto-fallend' | 'goto0after';
type Shape = 'inline' | 'loop' | 'child-unhandled' | 'child-handled' | 'double-error' | 'error-in-handler';

interface ErrSource { stmt: string; num: number; label: string }
const ERR_SOURCES: ErrSource[] = [
    { stmt: 'Err.Raise 5', num: 5, label: 'Raise5' },
    { stmt: 'z = 1 \\ zero', num: 11, label: 'Div0' },
    { stmt: 'z = CLng("abc")', num: 13, label: 'Mismatch' },
    { stmt: 'arr(99) = 1', num: 9, label: 'Subscript' },
];

interface Scenario {
    name: string;
    code: string;
    childCode?: string;
    expected: { trace: string[]; fatalNum: number | null };
}

const HEADER = 'Dim z As Variant\n    Dim zero As Long\n    Dim arr(2) As Long';

function modePrologue(mode: Mode): string {
    switch (mode) {
        case 'none': return '';
        case 'resumenext': return 'On Error Resume Next';
        case 'goto0after': return 'On Error Resume Next';
        default: return 'On Error GoTo H';
    }
}

function handlerBlock(mode: Mode): string {
    if (mode === 'goto-resumenext') return 'H:\n    Debug.Print "H" & Err.Number\n    Resume Next';
    if (mode === 'goto-fallend') return 'H:\n    Debug.Print "H" & Err.Number';
    return '';
}

/** インライン形状: S1 → (エラー位置 0..2 のいずれか) → プローブ → S2 → SF */
function buildInline(mode: Mode, errPos: number, src: ErrSource): Scenario {
    const lines: string[] = [];
    const trace: string[] = [];
    let fatalNum: number | null = null;
    const handled = mode === 'resumenext' || mode.startsWith('goto-');

    lines.push(modePrologue(mode));
    const logs = ['S1', 'S2', 'S3'];
    let aborted = false;
    for (let i = 0; i < logs.length; i++) {
        if (i === errPos) {
            if (mode === 'goto0after') lines.push('On Error GoTo 0');
            lines.push(src.stmt);
            lines.push('Debug.Print "N" & Err.Number');
            if (!aborted) {
                if (mode === 'resumenext') {
                    trace.push(`N${src.num}`); // Err は保持
                } else if (mode === 'goto-resumenext') {
                    trace.push(`H${src.num}`);
                    trace.push('N0'); // Resume で Err クリア
                } else if (mode === 'goto-fallend') {
                    trace.push(`H${src.num}`);
                    aborted = true; // ハンドラーが End Sub に落ちて正常終了
                } else {
                    fatalNum = src.num;
                    aborted = true;
                }
            }
        }
        lines.push(`Debug.Print "${logs[i]}"`);
        if (!aborted) trace.push(logs[i]);
    }
    const body = lines.filter(l => l !== '').map(l => '    ' + l).join('\n');
    const code = `Sub Main()\n    ${HEADER}\n${body}\n    Exit Sub\n${handlerBlock(mode)}\nEnd Sub`;
    return { name: `inline/${mode}/pos${errPos}/${src.label}`, code, expected: { trace, fatalNum } };
}

/** ループ形状: For 2 周の中でエラー。Resume Next 系では全周回が継続する */
function buildLoop(mode: Mode, src: ErrSource): Scenario {
    const trace: string[] = [];
    let fatalNum: number | null = null;
    const iterations = 2;
    let aborted = false;
    for (let it = 1; it <= iterations && !aborted; it++) {
        trace.push(`L${it}`);
        if (mode === 'resumenext') {
            trace.push(`N${src.num}`);
        } else if (mode === 'goto-resumenext') {
            trace.push(`H${src.num}`, 'N0');
        } else if (mode === 'goto-fallend') {
            trace.push(`H${src.num}`);
            aborted = true;
            break;
        } else {
            fatalNum = src.num;
            aborted = true;
            break;
        }
        trace.push(`E${it}`);
    }
    if (!aborted) trace.push('SF');

    const goto0 = mode === 'goto0after' ? '\n    On Error GoTo 0' : '';
    const code = `Sub Main()
    ${HEADER}
    ${modePrologue(mode)}${goto0}
    Dim i As Long
    For i = 1 To ${iterations}
        Debug.Print "L" & i
        ${src.stmt}
        Debug.Print "N" & Err.Number
        Debug.Print "E" & i
    Next i
    Debug.Print "SF"
    Exit Sub
${handlerBlock(mode)}
End Sub`;
    return { name: `loop/${mode}/${src.label}`, code, expected: { trace, fatalNum } };
}

/** 子プロシージャ形状: Main から Child を呼び、Child 内でエラー */
function buildChild(mode: Mode, childHandled: boolean, src: ErrSource): Scenario {
    const trace: string[] = ['S1'];
    let fatalNum: number | null = null;
    let aborted = false;

    trace.push('C1');
    if (childHandled) {
        // Child が自前の Resume Next で処理 → Child 正常終了。
        // End Sub は Err を自動クリアしない（§6.1.3.2）ため Main のプローブにも Err が見える
        trace.push('C2', `N${src.num}`, 'S2');
    } else {
        // Child 未処理 → エラーは Main の Call 文の位置で発生した扱い
        if (mode === 'resumenext') {
            trace.push(`N${src.num}`, 'S2');
        } else if (mode === 'goto-resumenext') {
            trace.push(`H${src.num}`, 'N0', 'S2');
        } else if (mode === 'goto-fallend') {
            trace.push(`H${src.num}`);
            aborted = true;
        } else {
            fatalNum = src.num;
            aborted = true;
        }
    }

    const goto0 = mode === 'goto0after' ? '\n    On Error GoTo 0' : '';
    const code = `Sub Main()
    ${HEADER}
    ${modePrologue(mode)}${goto0}
    Debug.Print "S1"
    Child
    Debug.Print "N" & Err.Number
    Debug.Print "S2"
    Exit Sub
${handlerBlock(mode)}
End Sub`;
    const childCode = childHandled
        ? `Sub Child()
    Dim z As Variant
    Dim zero As Long
    Dim arr(2) As Long
    On Error Resume Next
    Debug.Print "C1"
    ${src.stmt}
    Debug.Print "C2"
End Sub`
        : `Sub Child()
    Dim z As Variant
    Dim zero As Long
    Dim arr(2) As Long
    Debug.Print "C1"
    ${src.stmt}
    Debug.Print "C2"
End Sub`;
    return {
        name: `child/${mode}/${childHandled ? 'handled' : 'unhandled'}/${src.label}`,
        code, childCode, expected: { trace, fatalNum },
    };
}

/** 二重エラー形状: goto-resumenext でハンドラーが再武装され 2 回目も捕捉されること */
function buildDoubleError(src: ErrSource): Scenario {
    const trace = ['S1', `H${src.num}`, 'N0', 'S2', `H${src.num}`, 'N0', 'S3'];
    const code = `Sub Main()
    ${HEADER}
    On Error GoTo H
    Debug.Print "S1"
    ${src.stmt}
    Debug.Print "N" & Err.Number
    Debug.Print "S2"
    ${src.stmt}
    Debug.Print "N" & Err.Number
    Debug.Print "S3"
    Exit Sub
H:
    Debug.Print "H" & Err.Number
    Resume Next
End Sub`;
    return { name: `double-error/${src.label}`, code, expected: { trace, fatalNum: null } };
}

/** ハンドラー内エラー形状: アクティブなハンドラー内の再エラーは伝播して fatal */
function buildErrorInHandler(src: ErrSource): Scenario {
    const trace = ['S1', `H${src.num}`];
    const code = `Sub Main()
    ${HEADER}
    On Error GoTo H
    Debug.Print "S1"
    ${src.stmt}
    Debug.Print "S2"
    Exit Sub
H:
    Debug.Print "H" & Err.Number
    ${src.stmt}
    Debug.Print "HX"
End Sub`;
    return { name: `error-in-handler/${src.label}`, code, expected: { trace, fatalNum: src.num } };
}

// ---------------------------------------------------------------------------
// 実行
// ---------------------------------------------------------------------------
const MODES: Mode[] = ['none', 'resumenext', 'goto-resumenext', 'goto-fallend', 'goto0after'];
const scenarios: Scenario[] = [];

for (const mode of MODES) {
    for (const src of ERR_SOURCES) {
        for (let pos = 0; pos < 3; pos++) scenarios.push(buildInline(mode, pos, src));
        scenarios.push(buildLoop(mode, src));
        scenarios.push(buildChild(mode, false, src));
        scenarios.push(buildChild(mode, true, src));
    }
}
for (const src of ERR_SOURCES) {
    scenarios.push(buildDoubleError(src));
    scenarios.push(buildErrorInHandler(src));
}

let failures = 0;
console.log(`On Error ストレス: ${scenarios.length} シナリオ`);

for (const sc of scenarios) {
    const trace: string[] = [];
    let fatal: number | null = null;
    try {
        const full = sc.childCode ? `${sc.code}\n\n${sc.childCode}` : sc.code;
        const ev = evalVBASingle(full, { onPrint: (s: string) => trace.push(s) });
        ev.callProcedure('Main', []);
    } catch (e: any) {
        const m = /error '(\d+)'/.exec(String(e?.message ?? e));
        fatal = m ? Number(m[1]) : -1;
        if (fatal === -1) {
            failures++;
            console.log(`\n❌ ${sc.name} — VBA エラー以外の例外: ${String(e?.message ?? e).slice(0, 140)}`);
            continue;
        }
    }
    const traceOk = JSON.stringify(trace) === JSON.stringify(sc.expected.trace);
    const fatalOk = fatal === sc.expected.fatalNum;
    if (!traceOk || !fatalOk) {
        failures++;
        console.log(`\n❌ ${sc.name}`);
        if (!traceOk) console.log(`   trace 期待 ${JSON.stringify(sc.expected.trace)}\n   trace 実際 ${JSON.stringify(trace)}`);
        if (!fatalOk) console.log(`   fatal 期待 ${sc.expected.fatalNum} / 実際 ${fatal}`);
    }
}

console.log(`\n${scenarios.length} シナリオ中 不一致 ${failures} 件`);
if (failures > 0) {
    process.exitCode = 1;
} else {
    console.log('✅ 全シナリオ一致');
}
