/**
 * サンプルソースに意図的なエラーを埋め込み、エラー報告（行番号・エラー番号・モジュール）が
 * 正しいかを検証する。
 *
 * テスト対象エラー種別:
 *  1. ゼロ除算 (Error 11)
 *  2. 型の不一致 (Error 13)
 *  3. Option Explicit 違反 (Error 1)
 *  4. 添字が有効範囲外 (Error 9)
 *  5. プロシージャが未定義 (Error 35)
 *  6. 深い呼び出しのスタックトレース
 */
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

interface VbaError {
    message: string;
    number: number;
    vbaLine?: number;
    vbaModule?: string | null;
    vbaStack?: Array<{ name: string; moduleName: string; line: number }>;
}

function catchError(
    modules: Array<{ name: string; code: string }>,
    entry: string,
): VbaError {
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

// ---------------------------------------------------------------------------
// Test 1: ゼロ除算 (Error 11) — TaskScheduler_Core 風の計算関数にエラーを埋め込む
// ---------------------------------------------------------------------------
{
    const code = [
        'Function CalcDuration(numDays As Long, capacity As Double) As Double', // 1
        '    If capacity = 0 Then',                                               // 2
        '        CalcDuration = numDays / capacity',                              // 3  ← ゼロ除算
        '    Else',                                                               // 4
        '        CalcDuration = numDays / capacity',                              // 5
        '    End If',                                                             // 6
        'End Function',                                                           // 7
        '',                                                                       // 8
        'Sub RunCalc()',                                                          // 9
        '    Dim result As Double',                                               // 10
        '    result = CalcDuration(10, 0)',                                       // 11  ← 呼び出し行
        'End Sub',                                                                // 12
    ].join('\n');

    const e = catchError([{ name: 'CalcModule', code }], 'RunCalc');
    assert.strictEqual(e.number, 11, 'Test1: ゼロ除算は Error 11');
    assert.strictEqual(e.vbaLine, 3, 'Test1: エラー行は 3');
    assert.strictEqual(e.vbaModule, 'CalcModule', 'Test1: モジュール名は CalcModule');

    const stack = e.vbaStack!;
    assert.ok(Array.isArray(stack) && stack.length === 2, 'Test1: スタックは 2 フレーム');
    assert.strictEqual(stack[0].name, 'CalcDuration', 'Test1: フレーム0 = CalcDuration');
    assert.strictEqual(stack[1].name, 'RunCalc', 'Test1: フレーム1 = RunCalc');
    assert.strictEqual(stack[0].line, 11, 'Test1: CalcDuration は RunCalc の 11 行目から呼ばれた');
    console.log('[PASS] Test 1: ゼロ除算の行番号とスタックトレース');
}

// ---------------------------------------------------------------------------
// Test 2: 型の不一致 (Error 13) — 文字列を数値変数へ代入
// ---------------------------------------------------------------------------
{
    const code = [
        'Sub ProcessData()',                           // 1
        '    Dim level As Long',                       // 2
        '    Dim rawValue As String',                  // 3
        '    rawValue = "abc"',                        // 4
        '    level = CLng(rawValue)',                  // 5  ← 型の不一致
        '    level = level + 1',                       // 6
        'End Sub',                                     // 7
    ].join('\n');

    const e = catchError([{ name: 'DataModule', code }], 'ProcessData');
    assert.strictEqual(e.number, 13, 'Test2: 型不一致は Error 13');
    assert.strictEqual(e.vbaLine, 5, 'Test2: エラー行は 5');
    assert.strictEqual(e.vbaModule, 'DataModule', 'Test2: モジュール名は DataModule');
    console.log('[PASS] Test 2: 型の不一致の行番号');
}

// ---------------------------------------------------------------------------
// Test 3: Option Explicit 違反 (Error 1) — 未宣言変数の使用
// ---------------------------------------------------------------------------
{
    const code = [
        'Option Explicit',                             // 1
        '',                                            // 2
        'Sub InitConfig()',                            // 3
        '    Dim rowStart As Long',                    // 4
        '    rowStart = 10',                           // 5
        '    totalRows = rowStart + 5',                // 6  ← 未宣言 totalRows
        'End Sub',                                     // 7
    ].join('\n');

    const e = catchError([{ name: 'ConfigModule', code }], 'InitConfig');
    assert.strictEqual(e.number, 1, 'Test3: OE 違反は Error 1');
    assert.strictEqual(e.vbaLine, 6, 'Test3: エラー行は 6');
    assert.strictEqual(e.vbaModule, 'ConfigModule', 'Test3: モジュール名は ConfigModule');
    assert.ok(e.message.includes('totalrows'), `Test3: メッセージに未宣言名を含む: ${e.message}`);
    console.log('[PASS] Test 3: Option Explicit 違反の行番号');
}

// ---------------------------------------------------------------------------
// Test 4: オーバーフロー (Error 6) — Byte 型に範囲外の値を代入
// ---------------------------------------------------------------------------
{
    const code = [
        'Sub ValidateCapacity()',                       // 1
        '    Dim maxLoad As Long',                      // 2
        '    maxLoad = 200',                            // 3
        '    Dim compactLoad As Byte',                  // 4
        '    compactLoad = maxLoad + 100',              // 5  ← 300 は Byte(0-255) を超えてオーバーフロー
        'End Sub',                                      // 6
    ].join('\n');

    const e = catchError([{ name: 'ValidModule', code }], 'ValidateCapacity');
    assert.strictEqual(e.number, 6, 'Test4: オーバーフローは Error 6');
    assert.strictEqual(e.vbaLine, 5, 'Test4: エラー行は 5');
    assert.strictEqual(e.vbaModule, 'ValidModule', 'Test4: モジュール名は ValidModule');
    console.log('[PASS] Test 4: オーバーフローの行番号');
}

// ---------------------------------------------------------------------------
// Test 5: 未定義プロシージャ (Error 35) — 存在しない Sub の呼び出し
// ---------------------------------------------------------------------------
{
    const code = [
        'Sub ScheduleTask()',                          // 1
        '    Dim taskRow As Long',                     // 2
        '    taskRow = 1',                             // 3
        '    Call UpdateCalendar(taskRow)',            // 4  ← 未定義 UpdateCalendar
        'End Sub',                                     // 5
    ].join('\n');

    const e = catchError([{ name: 'ScheduleModule', code }], 'ScheduleTask');
    assert.strictEqual(e.number, 35, 'Test5: 未定義プロシージャは Error 35');
    assert.strictEqual(e.vbaLine, 4, 'Test5: エラー行は 4');
    assert.strictEqual(e.vbaModule, 'ScheduleModule', 'Test5: モジュール名は ScheduleModule');
    console.log('[PASS] Test 5: 未定義プロシージャの行番号');
}

// ---------------------------------------------------------------------------
// Test 6: 3 階層の呼び出しスタックトレース — TaskScheduler 風の呼び出し構造
// ---------------------------------------------------------------------------
{
    const code = [
        'Sub RunScheduler()',               // 1
        '    Call ProcessLevel(2)',         // 2
        'End Sub',                          // 3
        '',                                 // 4
        'Sub ProcessLevel(lvl As Long)',    // 5
        '    Call AllocateDay(lvl)',        // 6
        'End Sub',                          // 7
        '',                                 // 8
        'Sub AllocateDay(lvl As Long)',     // 9
        '    Dim cap As Double',            // 10
        '    cap = 0',                      // 11
        '    Dim alloc As Double',          // 12
        '    alloc = 1 / cap',             // 13  ← ゼロ除算
        'End Sub',                          // 14
    ].join('\n');

    const e = catchError([{ name: 'Scheduler', code }], 'RunScheduler');
    assert.strictEqual(e.number, 11, 'Test6: ゼロ除算は Error 11');
    assert.strictEqual(e.vbaLine, 13, 'Test6: エラー行は 13');

    const stack = e.vbaStack!;
    assert.strictEqual(stack.length, 3, 'Test6: スタックは 3 フレーム');
    assert.strictEqual(stack[0].name, 'AllocateDay', 'Test6: フレーム0 = AllocateDay');
    assert.strictEqual(stack[1].name, 'ProcessLevel', 'Test6: フレーム1 = ProcessLevel');
    assert.strictEqual(stack[2].name, 'RunScheduler', 'Test6: フレーム2 = RunScheduler');
    assert.strictEqual(stack[0].line, 6, 'Test6: AllocateDay は ProcessLevel の 6 行目から呼ばれた');
    assert.strictEqual(stack[1].line, 2, 'Test6: ProcessLevel は RunScheduler の 2 行目から呼ばれた');
    assert.strictEqual(stack[2].line, 0, 'Test6: RunScheduler はエントリポイント');
    console.log('[PASS] Test 6: 3 階層スタックトレースの行番号');
}

// ---------------------------------------------------------------------------
// Test 7: クロスモジュール — エラーが正しいモジュール名で報告される
// ---------------------------------------------------------------------------
{
    const coreCode = [
        'Function CalcBaseStart(level As Long, parentFinish As Long, parentAlloc As Double) As Long', // 1
        '    If level > 1 Then',                                                                       // 2
        '        Dim x As Long',                                                                       // 3
        '        x = 1 / (parentFinish - parentFinish)',                                               // 4  ← ゼロ除算
        '    End If',                                                                                  // 5
        '    CalcBaseStart = 1',                                                                       // 6
        'End Function',                                                                                // 7
    ].join('\n');

    const runnerCode = [
        'Sub RunTest()',                                    // 1
        '    Dim result As Long',                           // 2
        '    result = CalcBaseStart(2, 5, 0.3)',           // 3
        'End Sub',                                          // 4
    ].join('\n');

    const e = catchError(
        [{ name: 'CoreModule', code: coreCode }, { name: 'RunnerModule', code: runnerCode }],
        'RunTest',
    );
    assert.strictEqual(e.number, 11, 'Test7: ゼロ除算は Error 11');
    assert.strictEqual(e.vbaLine, 4, 'Test7: エラー行は CoreModule の 4');
    assert.strictEqual(e.vbaModule, 'CoreModule', 'Test7: エラーモジュールは CoreModule');

    const stack = e.vbaStack!;
    assert.strictEqual(stack[0].moduleName, 'CoreModule', 'Test7: フレーム0 のモジュールは CoreModule');
    assert.strictEqual(stack[1].moduleName, 'RunnerModule', 'Test7: フレーム1 のモジュールは RunnerModule');
    console.log('[PASS] Test 7: クロスモジュールのエラーモジュール名');
}

console.log('\n✅ error-line-reporting: 全テスト通過');
