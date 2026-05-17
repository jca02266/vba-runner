/**
 * String と数値の暗黙変換 (§5.6.9.3 周辺)
 *
 * VBA の演算子は型ごとに異なる暗黙変換ルールを持つ:
 *
 *   - `+`  : 両方が String → 連結。それ以外（数値混在含む）→ 数値加算
 *           （文字列が数値化できない場合は Type mismatch エラー 13）
 *   - `&`  : 常に文字列連結（数値は文字列化）
 *   - `-` `*` `/` `\` `Mod` `^` : 常に数値演算（変換失敗で Error 13）
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

function expectError(fn: () => any, errorNumber: number, label: string) {
    let caught: any = null;
    try { fn(); } catch (e: any) { caught = e; }
    if (!caught) {
        console.error(`[FAIL] ${label} - エラーが発生しなかった`);
        throw new Error('Assertion Failed');
    }
    const num = caught.number ?? (caught.message && caught.message.match(/error (\d+)/i)?.[1]);
    if (Number(num) !== errorNumber) {
        console.error(`[FAIL] ${label} - 期待エラー番号 ${errorNumber} だが、得たもの: ${JSON.stringify(caught)}`);
        throw new Error('Assertion Failed');
    }
}

console.log('--- Starting String/Numeric Coercion Tests ---');

// =============================================================================
// 1. `+` 演算子 — 数値混在は数値加算、両方 String は連結
// =============================================================================
{
    const code = `
        Function S1() : S1 = "5" + 3 : End Function
        Function S2() : S2 = 3 + "5" : End Function
        Function S3() : S3 = "5" + "3" : End Function
        Function S4() : S4 = "3.14" + 1 : End Function
        Function S5() : S5 = "-5" + 3 : End Function
        Function S6() : S6 = "  10  " + 1 : End Function
    `;
    assert.strictEqual(runFunc(code, 'S1'), 8, '"5" + 3 = 8 (数値混在)');
    assert.strictEqual(runFunc(code, 'S2'), 8, '3 + "5" = 8');
    assert.strictEqual(runFunc(code, 'S3'), '53', '"5" + "3" = "53" (両方 String は連結)');
    // 浮動小数点演算の特性で 4.140000000000001 になる（IEEE 754 倍精度）
    const s4 = runFunc(code, 'S4') as number;
    assert.ok(Math.abs(s4 - 4.14) < 1e-10, `"3.14" + 1 ≈ 4.14 (got ${s4})`);
    assert.strictEqual(runFunc(code, 'S5'), -2, '"-5" + 3 = -2');
    assert.strictEqual(runFunc(code, 'S6'), 11, '前後空白も許容 "  10  " + 1 = 11');
    console.log('[PASS] + 演算子の文字列/数値混在ルール');
}

// =============================================================================
// 2. 算術演算子 (-, *, /, \, Mod, ^) — 常に数値変換
// =============================================================================
{
    const code = `
        Function Sub1() : Sub1 = "10" - "3" : End Function
        Function Mul() : Mul = "5" * 2 : End Function
        Function Div() : Div = "100" / "4" : End Function
        Function IDiv() : IDiv = "7" \\ 2 : End Function
        Function Mo() : Mo = "7" Mod 3 : End Function
        Function Pow() : Pow = "2" ^ "3" : End Function
    `;
    assert.strictEqual(runFunc(code, 'Sub1'), 7, '"10" - "3" = 7');
    assert.strictEqual(runFunc(code, 'Mul'), 10, '"5" * 2 = 10');
    assert.strictEqual(runFunc(code, 'Div'), 25, '"100" / "4" = 25');
    assert.strictEqual(runFunc(code, 'IDiv'), 3, '"7" \\ 2 = 3');
    assert.strictEqual(runFunc(code, 'Mo'), 1, '"7" Mod 3 = 1');
    assert.strictEqual(runFunc(code, 'Pow'), 8, '"2" ^ "3" = 8');
    console.log('[PASS] -, *, /, \\, Mod, ^ の文字列→数値変換');
}

// =============================================================================
// 3. 数値変換できない文字列 → Type mismatch (Error 13)
// =============================================================================
{
    const code = `
        Function E1() : E1 = "abc" + 1 : End Function
        Function E2() : E2 = "abc" * 2 : End Function
        Function E3() : E3 = "abc" - 1 : End Function
        Function E4() : E4 = "abc" / 2 : End Function
    `;
    const ev = evalVBA(code);
    expectError(() => ev.callProcedure('E1', []), 13, '"abc" + 1 → Type mismatch');
    expectError(() => ev.callProcedure('E2', []), 13, '"abc" * 2 → Type mismatch');
    expectError(() => ev.callProcedure('E3', []), 13, '"abc" - 1 → Type mismatch');
    expectError(() => ev.callProcedure('E4', []), 13, '"abc" / 2 → Type mismatch');
    console.log('[PASS] 変換不能な文字列での Type mismatch');
}

// =============================================================================
// 4. `&` 演算子 — 常に文字列連結（数値も文字列化）
// =============================================================================
{
    const code = `
        Function C1() : C1 = "5" & 3 : End Function
        Function C2() : C2 = 5 & "3" : End Function
        Function C3() : C3 = 5 & 3 : End Function
        Function C4() : C4 = "abc" & 1 : End Function
        Function C5() : C5 = 1 & 2 & 3 : End Function
    `;
    assert.strictEqual(runFunc(code, 'C1'), '53', '"5" & 3 = "53"');
    assert.strictEqual(runFunc(code, 'C2'), '53', '5 & "3" = "53"');
    assert.strictEqual(runFunc(code, 'C3'), '53', '5 & 3 = "53"（両方数値でも連結）');
    assert.strictEqual(runFunc(code, 'C4'), 'abc1', '"abc" & 1 = "abc1" (連結は数値化不能でも OK)');
    assert.strictEqual(runFunc(code, 'C5'), '123', '連続する & 演算子');
    console.log('[PASS] & 演算子は常に連結');
}

// =============================================================================
// 5. VBA ベテランが書きそうな代表コード — CSV / 文字列パース
// =============================================================================
// セルから読んだ値が文字列だが、数値として扱いたい場面
{
    const code = `
        ' セルの値が "100" や "200.5" のような数値文字列で来る想定
        Function ParseCsvLine(line As String) As Double
            Dim parts() As String
            parts = Split(line, ",")
            Dim total As Double
            Dim i As Long
            For i = 0 To UBound(parts)
                total = total + parts(i)  ' 文字列 + 累計 (数値) → 数値加算
            Next i
            ParseCsvLine = total
        End Function

        ' 範囲合計（文字列っぽい数値混在）
        Function MixedSum() As Double
            MixedSum = "10" + 20 + "30.5" + 40
        End Function

        ' 数値文字列 → 連結したいときは & を明示
        Function MakeId(prefix As String, num As Long) As String
            MakeId = prefix & "-" & num
        End Function
    `;
    assert.strictEqual(runFunc(code, 'ParseCsvLine', ['100,200,50']), 350, 'CSV 行の文字列合計');
    assert.strictEqual(runFunc(code, 'ParseCsvLine', ['1.5,2.5,3.0']), 7, '小数点の文字列合計');
    assert.strictEqual(runFunc(code, 'MixedSum'), 100.5, '文字列/数値混在の + 計算');
    assert.strictEqual(runFunc(code, 'MakeId', ['ORD', 42]), 'ORD-42', '& で ID 生成');
    console.log('[PASS] VBA ベテランパターン: CSV / 文字列パース');
}

// =============================================================================
// 6. VBA ベテランが書きそうな代表コード — 入力値の数値判定
// =============================================================================
// InputBox / Range.Value から文字列が来る想定
{
    const code = `
        ' VBA 慣用句: IsNumeric で数値判定 → そのまま算術に使える
        Function SafeAdd(s As String) As Double
            If IsNumeric(s) Then
                SafeAdd = s + 100  ' 文字列でも IsNumeric なら数値演算 OK
            Else
                SafeAdd = -1
            End If
        End Function

        ' Val 関数で数値部分のみ抽出
        Function ExtractNum(s As String) As Double
            ExtractNum = Val(s) * 2
        End Function
    `;
    assert.strictEqual(runFunc(code, 'SafeAdd', ['50']), 150, 'IsNumeric → 文字列+数値');
    assert.strictEqual(runFunc(code, 'SafeAdd', ['abc']), -1, 'IsNumeric=False → ガード');
    assert.strictEqual(runFunc(code, 'ExtractNum', ['123abc']), 246, 'Val で数値抽出 → 乗算');
    console.log('[PASS] VBA ベテランパターン: 入力値の数値判定');
}

console.log('\n✅ String/Numeric 暗黙変換: 全テスト通過');
