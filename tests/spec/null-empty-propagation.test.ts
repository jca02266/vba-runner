/**
 * Empty / Null / Nothing / Missing の演算子伝播ルール (§5.6.9 周辺)
 *
 * VBA 仕様:
 *   - Null: 任意の算術・比較演算で Null が伝播（例外: 文字列連結 & では空文字扱い）
 *   - Empty: 数値 0 または "" と同等として扱われる
 *   - Nothing: オブジェクト参照のみ。算術演算では Type mismatch エラー
 *   - Missing: 未指定の Optional 引数（IsMissing で検出）
 */
import { vbaNull, vbaEmpty } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

console.log('--- Starting Null/Empty Propagation Tests ---');

// =============================================================================
// 1. Null の算術伝播 — VBA 仕様: 算術演算で Null は必ず Null を返す
// =============================================================================
{
    const code = `
        Function NullPlus5() : NullPlus5 = Null + 5 : End Function
        Function FivePlusNull() : FivePlusNull = 5 + Null : End Function
        Function NullTimes() : NullTimes = Null * 10 : End Function
        Function NullDiv() : NullDiv = Null / 2 : End Function
        Function NullMinus() : NullMinus = Null - 7 : End Function
        Function NullMod() : NullMod = Null Mod 3 : End Function
        Function NullPow() : NullPow = Null ^ 2 : End Function
        Function NullIntDiv() : NullIntDiv = Null \\ 2 : End Function
        Function NullNeg() : NullNeg = -Null : End Function
    `;
    assert.strictEqual(runFunc(code, 'NullPlus5'), vbaNull, 'Null + 5 = Null');
    assert.strictEqual(runFunc(code, 'FivePlusNull'), vbaNull, '5 + Null = Null');
    assert.strictEqual(runFunc(code, 'NullTimes'), vbaNull, 'Null * 10 = Null');
    assert.strictEqual(runFunc(code, 'NullDiv'), vbaNull, 'Null / 2 = Null');
    assert.strictEqual(runFunc(code, 'NullMinus'), vbaNull, 'Null - 7 = Null');
    assert.strictEqual(runFunc(code, 'NullMod'), vbaNull, 'Null Mod 3 = Null');
    assert.strictEqual(runFunc(code, 'NullPow'), vbaNull, 'Null ^ 2 = Null');
    assert.strictEqual(runFunc(code, 'NullIntDiv'), vbaNull, 'Null \\ 2 = Null');
    assert.strictEqual(runFunc(code, 'NullNeg'), vbaNull, '-Null = Null');
    console.log('[PASS] Null 算術伝播');
}

// =============================================================================
// 2. Null と文字列演算
// =============================================================================
{
    const code = `
        ' & は VBA 仕様の特例: Null を "" として扱う
        Function NullConcatStr() : NullConcatStr = Null & "abc" : End Function
        Function StrConcatNull() : StrConcatNull = "abc" & Null : End Function
        Function NullConcatNull() : NullConcatNull = Null & Null : End Function

        ' + は通常の算術扱い: Null + 文字列 = Null
        Function NullPlusStr() : NullPlusStr = Null + "abc" : End Function
    `;
    assert.strictEqual(runFunc(code, 'NullConcatStr'), 'abc', 'Null & "abc" = "abc"');
    assert.strictEqual(runFunc(code, 'StrConcatNull'), 'abc', '"abc" & Null = "abc"');
    assert.strictEqual(runFunc(code, 'NullConcatNull'), '', 'Null & Null = ""');
    assert.strictEqual(runFunc(code, 'NullPlusStr'), vbaNull, 'Null + "abc" = Null');
    console.log('[PASS] Null と文字列演算');
}

// =============================================================================
// 3. Null と比較演算
// =============================================================================
// Null を含む比較は VBA 仕様で常に Null を返す（True/False ではない）
{
    const code = `
        Function NullEqNull() : NullEqNull = (Null = Null) : End Function
        Function NullEq5() : NullEq5 = (Null = 5) : End Function
        Function NullNeq5() : NullNeq5 = (Null <> 5) : End Function
        Function NullLt5() : NullLt5 = (Null < 5) : End Function
        Function NullGt5() : NullGt5 = (Null > 5) : End Function
    `;
    assert.strictEqual(runFunc(code, 'NullEqNull'), vbaNull, 'Null = Null → Null');
    assert.strictEqual(runFunc(code, 'NullEq5'), vbaNull, 'Null = 5 → Null');
    assert.strictEqual(runFunc(code, 'NullNeq5'), vbaNull, 'Null <> 5 → Null');
    assert.strictEqual(runFunc(code, 'NullLt5'), vbaNull, 'Null < 5 → Null');
    assert.strictEqual(runFunc(code, 'NullGt5'), vbaNull, 'Null > 5 → Null');
    console.log('[PASS] Null 比較演算');
}

// =============================================================================
// 4. Empty の算術伝播 — VBA 仕様: Empty は数値文脈では 0 として扱う
// =============================================================================
{
    const code = `
        Function EmptyPlus5() : EmptyPlus5 = Empty + 5 : End Function
        Function EmptyTimes() : EmptyTimes = Empty * 7 : End Function
        Function EmptyMinus() : EmptyMinus = Empty - 3 : End Function
        Function EmptyDiv() : EmptyDiv = Empty / 4 : End Function

        ' 両方 Empty の場合も数値 0 として処理
        Function EmptyPlusEmpty() : EmptyPlusEmpty = Empty + Empty : End Function
    `;
    assert.strictEqual(runFunc(code, 'EmptyPlus5'), 5, 'Empty + 5 = 5');
    assert.strictEqual(runFunc(code, 'EmptyTimes'), 0, 'Empty * 7 = 0');
    assert.strictEqual(runFunc(code, 'EmptyMinus'), -3, 'Empty - 3 = -3');
    assert.strictEqual(runFunc(code, 'EmptyDiv'), 0, 'Empty / 4 = 0');
    assert.strictEqual(runFunc(code, 'EmptyPlusEmpty'), 0, 'Empty + Empty = 0');
    console.log('[PASS] Empty 算術伝播');
}

// =============================================================================
// 5. Empty と文字列演算 — VBA 仕様: 文字列文脈では "" として扱う
// =============================================================================
{
    const code = `
        Function EmptyConcat() : EmptyConcat = Empty & "abc" : End Function
        Function StrConcatEmpty() : StrConcatEmpty = "abc" & Empty : End Function
        Function EmptyPlusEmptyStr() : EmptyPlusEmptyStr = Empty + "" : End Function
        Function EmptyConcatEmpty() : EmptyConcatEmpty = Empty & Empty : End Function
    `;
    assert.strictEqual(runFunc(code, 'EmptyConcat'), 'abc', 'Empty & "abc" = "abc"');
    assert.strictEqual(runFunc(code, 'StrConcatEmpty'), 'abc', '"abc" & Empty = "abc"');
    assert.strictEqual(runFunc(code, 'EmptyPlusEmptyStr'), '', 'Empty + "" = ""');
    assert.strictEqual(runFunc(code, 'EmptyConcatEmpty'), '', 'Empty & Empty = ""');
    console.log('[PASS] Empty 文字列演算');
}

// =============================================================================
// 6. Empty と比較演算 — VBA 仕様: 0 / "" と等価
// =============================================================================
{
    const code = `
        Function EmptyEq0() : EmptyEq0 = (Empty = 0) : End Function
        Function EmptyEqStr() : EmptyEqStr = (Empty = "") : End Function
        Function EmptyEqEmpty() : EmptyEqEmpty = (Empty = Empty) : End Function
        Function EmptyLt5() : EmptyLt5 = (Empty < 5) : End Function
    `;
    assert.isTrue(runFunc(code, 'EmptyEq0'), 'Empty = 0 → True');
    assert.isTrue(runFunc(code, 'EmptyEqStr'), 'Empty = "" → True');
    assert.isTrue(runFunc(code, 'EmptyEqEmpty'), 'Empty = Empty → True');
    assert.isTrue(runFunc(code, 'EmptyLt5'), 'Empty < 5 → True (Empty を 0 として比較)');
    console.log('[PASS] Empty 比較演算');
}

// =============================================================================
// 7. VBA ベテランが書きそうな代表コード — Null 安全パターン
// =============================================================================
// 実務でよくある「DB / レコードセットから NULL が返ってくる」想定のコード
{
    const code = `
        ' レコードから取得した値が Null かもしれない時の典型的なガード
        Function SafeStr(v As Variant) As String
            If IsNull(v) Then
                SafeStr = ""
            Else
                SafeStr = CStr(v)
            End If
        End Function

        Function SafeNum(v As Variant) As Long
            If IsNull(v) Then
                SafeNum = 0
            Else
                SafeNum = CLng(v)
            End If
        End Function

        ' Null 伝播を利用したショートサーキット
        Function NullPropagation(a As Variant, b As Variant) As Variant
            NullPropagation = a + b  ' どちらかが Null なら結果は Null
        End Function

        ' 文字列連結では Null を許容（安全に concat）
        Function SafeConcat(prefix As String, v As Variant) As String
            SafeConcat = prefix & v   ' v が Null なら prefix & "" = prefix
        End Function
    `;
    const ev = evalVBA(code);

    assert.strictEqual(ev.callProcedure('SafeStr', [vbaNull]), '', 'SafeStr(Null) = ""');
    assert.strictEqual(ev.callProcedure('SafeStr', ['hello']), 'hello', 'SafeStr("hello") = "hello"');
    assert.strictEqual(ev.callProcedure('SafeNum', [vbaNull]), 0, 'SafeNum(Null) = 0');
    assert.strictEqual(ev.callProcedure('SafeNum', [42]), 42, 'SafeNum(42) = 42');

    assert.strictEqual(ev.callProcedure('NullPropagation', [vbaNull, 5]), vbaNull, '5 + Null = Null');
    assert.strictEqual(ev.callProcedure('NullPropagation', [3, 5]), 8, '3 + 5 = 8');

    assert.strictEqual(ev.callProcedure('SafeConcat', ['Name: ', vbaNull]), 'Name: ', '"Name: " & Null = "Name: "');
    assert.strictEqual(ev.callProcedure('SafeConcat', ['Name: ', 'Alice']), 'Name: Alice', '"Name: " & "Alice"');
    console.log('[PASS] VBA ベテランパターン: Null 安全コード');
}

// =============================================================================
// 8. VBA ベテランが書きそうな代表コード — 未初期化変数の Empty 利用
// =============================================================================
// 集計時に変数を宣言せず加算するパターン (初回は Empty + N = N で動く)
{
    const code = `
        Function SumLoop() As Long
            Dim total As Variant   ' 未初期化 → Empty
            Dim i As Long
            For i = 1 To 5
                total = total + i  ' 初回: Empty + 1 = 1
            Next i
            SumLoop = total
        End Function

        ' 未初期化文字列連結
        Function ConcatLoop() As String
            Dim result As Variant  ' Empty
            Dim i As Long
            For i = 1 To 3
                result = result & "[" & i & "]"  ' 初回: Empty & "[1]" = "[1]"
            Next i
            ConcatLoop = result
        End Function
    `;
    assert.strictEqual(runFunc(code, 'SumLoop'), 15, 'Empty 初期化での合計: 1+2+3+4+5=15');
    assert.strictEqual(runFunc(code, 'ConcatLoop'), '[1][2][3]', 'Empty 初期化での連結');
    console.log('[PASS] VBA ベテランパターン: 未初期化 Empty 活用');
}

// =============================================================================
// 9. Nothing - オブジェクト参照のみ（演算では Type mismatch）
// =============================================================================
// Nothing は基本的に演算対象にできない（仕様通り）
// ここでは IsNothing 相当の Is 演算子の挙動を検証
{
    const code = `
        Function IsObjNothing() As Boolean
            Dim o As Object
            IsObjNothing = (o Is Nothing)
        End Function

        Function NothingAfterSet() As Boolean
            Dim o As Object
            Set o = CreateObject("Scripting.Dictionary")
            Set o = Nothing
            NothingAfterSet = (o Is Nothing)
        End Function
    `;
    assert.isTrue(runFunc(code, 'IsObjNothing'), 'Dim o As Object → o Is Nothing = True');
    assert.isTrue(runFunc(code, 'NothingAfterSet'), 'Set o = Nothing → o Is Nothing = True');
    console.log('[PASS] Nothing オブジェクト参照');
}

console.log('\n✅ Null/Empty/Nothing/Missing 伝播ルール: 全テスト通過');
