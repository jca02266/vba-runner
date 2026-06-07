/**
 * 実行時エラーの行番号・モジュール名・スタックトレースのテスト。
 *
 * - VBA エラーオブジェクトは throwVbaError() で number/vbaLine/vbaModule/vbaStack を持つ。
 * - Option Explicit 違反はプロシージャ実行直前の AST スキャンで検出し、
 *   違反識別子の「最初の出現行」をエラー行として報告する。
 * - vbaStack は呼び出しの深い順（エラー発生プロシージャが先頭）に並ぶ。
 *   各フレームの line は「そのフレームが呼び出し元（次のフレーム）の何行目から呼ばれたか」。
 *   エントリポイント（最外フレーム）は呼び出し元がないので line=0。
 */
import { evalVBAModules } from '../../test-libs/test-runner';
import { assert } from '../../test-libs/test-runner';

interface VbaError {
    message: string;
    number: number;
    vbaLine?: number;
    vbaModule?: string | null;
    vbaStack?: Array<{ name: string; moduleName: string; line: number }>;
}

/** evalVBAModules または callProcedure で throw された VbaError を返す（throw されなければ失敗）。 */
function catchError(modules: Array<{ name: string; code: string }>, entry: string): VbaError {
    let ev: ReturnType<typeof evalVBAModules>;
    try {
        ev = evalVBAModules(modules);
    } catch (e: any) {
        return e as VbaError; // Option Explicit 等のコンパイルエラーは Pass2 で即時 throw
    }
    try {
        ev.callProcedure(entry, []);
    } catch (e: any) {
        return e as VbaError;
    }
    assert.fail(`expected ${entry} to throw a runtime error`);
    throw new Error('unreachable');
}

// --- Test 1: OE 違反の行番号は「最初の出現行」を指す ---
// 未宣言識別子 ThisWorkbook が 2 箇所に現れる。1 箇所目は呼び出し式の
// オブジェクト（callee.object）、2 箇所目はネストした MemberExpression。
// AST スキャン順に関係なく、最初の出現行（4 行目）を報告すること。
{
    const code = [
        'Option Explicit',                                   // 1
        'Private Function GetSheet(n As String) As Worksheet', // 2
        '    On Error Resume Next',                          // 3
        '    Set GetSheet = ThisWorkbook.Sheets(n)',         // 4  ← 最初の出現
        '    On Error GoTo 0',                               // 5
        '    If GetSheet Is Nothing Then',                   // 6
        '        Set GetSheet = ThisWorkbook.Sheets.Add',    // 7  ← ネストした参照
        '    End If',                                        // 8
        'End Function',                                      // 9
        'Public Sub Run()',                                  // 10
        '    Dim ws As Worksheet',                           // 11
        '    Set ws = GetSheet("Data")',                     // 12
        'End Sub',                                           // 13
    ].join('\n');

    const e = catchError([{ name: 'Mod1', code }], 'Run');
    assert.strictEqual(e.number, 1, 'OE 違反は Error 1');
    assert.strictEqual(e.vbaLine, 4, 'OE エラー行は最初の出現行(4)を指す');
    assert.strictEqual(e.vbaModule, 'Mod1', 'OE エラーのモジュールは違反プロシージャのモジュール');
    assert.ok(e.message.includes('thisworkbook'), `メッセージに未宣言名を含む: ${e.message}`);
    assert.ok(e.message.includes('(line 4)'), `メッセージに (line 4) を含む: ${e.message}`);
    console.log('[PASS] Test 1: OE 違反の行番号は最初の出現行を指す');
}

// --- Test 2: ランタイムエラー(ゼロ除算)のスタックトレース ---
// A(1行目から B を呼ぶ) → B(C を呼ぶ) → C(ゼロ除算) の 3 階層。
{
    const code = [
        'Sub A()',          // 1
        '    Call B()',     // 2
        'End Sub',          // 3
        'Sub B()',          // 4
        '    Call C()',     // 5
        'End Sub',          // 6
        'Sub C()',          // 7
        '    Dim x As Long',// 8
        '    x = 1 / 0',    // 9  ← エラー発生
        'End Sub',          // 10
    ].join('\n');

    const e = catchError([{ name: 'M', code }], 'A');
    assert.strictEqual(e.number, 11, 'ゼロ除算は Error 11');
    assert.strictEqual(e.vbaLine, 9, 'エラー発生行は 9');

    const stack = e.vbaStack!;
    assert.ok(Array.isArray(stack), 'vbaStack が配列で存在する');
    assert.strictEqual(stack.length, 3, 'スタックは 3 フレーム');

    // 深い順: C(発生), B, A
    assert.strictEqual(stack[0].name, 'C', 'フレーム0 = エラー発生プロシージャ C');
    assert.strictEqual(stack[1].name, 'B', 'フレーム1 = B');
    assert.strictEqual(stack[2].name, 'A', 'フレーム2 = エントリポイント A');

    assert.strictEqual(stack[0].moduleName, 'M', 'フレームのモジュール名は M');

    // line: そのフレームが呼び出し元の何行目から呼ばれたか
    assert.strictEqual(stack[0].line, 5, 'C は B の 5 行目から呼ばれた');
    assert.strictEqual(stack[1].line, 2, 'B は A の 2 行目から呼ばれた');
    assert.strictEqual(stack[2].line, 0, 'A はエントリポイント（呼び出し元なし）');
    console.log('[PASS] Test 2: ランタイムエラーのスタックトレース');
}

// --- Test 3: OE エラーは Pass2 コンパイル時に検出（スタックは空）---
// OE 違反は Pass2（resolveIdentifiers）で即時 throw するため、
// まだ callProcedure が実行されておらず vbaCallStack は空になる。
{
    const code = [
        'Option Explicit',           // 1
        'Public Sub Outer()',        // 2
        '    Inner',                 // 3
        'End Sub',                   // 4
        'Private Sub Inner()',       // 5
        '    x = undeclaredVar + 1', // 6  ← OE 違反
        'End Sub',                   // 7
    ].join('\n');

    const e = catchError([{ name: 'Mod', code }], 'Outer');
    assert.strictEqual(e.number, 1, 'OE 違反は Error 1');
    assert.strictEqual(e.vbaLine, 6, 'OE エラー行は違反行(6)');
    assert.strictEqual(e.vbaModule, 'Mod', 'OE エラーのモジュール');
    // Pass2 コンパイル時検出のためスタックは空
    assert.strictEqual((e.vbaStack ?? []).length, 0, 'コンパイル時検出なのでスタックは空');
    console.log('[PASS] Test 3: OE エラーは Pass2 コンパイル時に検出（スタックは空）');
}

// --- Test 4: クロスモジュール（モジュール修飾）呼び出しのスタックトレース ---
// ModA.Start が ModB.DoWork をモジュール修飾で呼ぶ。スタックの各フレームが
// それぞれのモジュール名を保持していること、および呼び出し先が投げた
// ランタイムエラーが握りつぶされず正しく伝播することを確認する。
{
    const modA = [
        'Public Sub Start()',     // 1
        '    ModB.DoWork',        // 2  ← モジュール修飾呼び出し
        'End Sub',                // 3
    ].join('\n');
    const modB = [
        'Public Sub DoWork()',    // 1
        '    Dim y As Long',      // 2
        '    y = 5 / 0',          // 3  ← エラー
        'End Sub',                // 4
    ].join('\n');

    const e = catchError([{ name: 'ModA', code: modA }, { name: 'ModB', code: modB }], 'Start');
    assert.strictEqual(e.number, 11, 'ゼロ除算は Error 11');
    assert.strictEqual(e.vbaModule, 'ModB', 'エラーは ModB で発生');
    assert.strictEqual(e.vbaLine, 3, 'ModB の 3 行目');

    const stack = e.vbaStack!;
    assert.strictEqual(stack.length, 2, 'スタックは 2 フレーム');
    assert.strictEqual(stack[0].name, 'DoWork', 'フレーム0 = DoWork');
    assert.strictEqual(stack[0].moduleName, 'ModB', 'DoWork は ModB');
    assert.strictEqual(stack[1].name, 'Start', 'フレーム1 = Start');
    assert.strictEqual(stack[1].moduleName, 'ModA', 'Start は ModA');
    assert.strictEqual(stack[0].line, 2, 'DoWork は Start の 2 行目から呼ばれた');
    assert.strictEqual(stack[1].line, 0, 'Start はエントリポイント');
    console.log('[PASS] Test 4: クロスモジュール呼び出しのスタックトレース');
}

console.log('\n✅ runtime-error-trace: 全テスト通過');
