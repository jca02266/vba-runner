/**
 * For...Next のループ変数と Next 後の変数が不一致の場合に
 * コンパイルエラーを生成するテスト
 *
 * VBA 仕様: `For i = 0 To 10 : Next j` はコンパイルエラー
 * 「Next で指定された変数の参照が不正です」
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function expectCompileError(code: string, label: string): void {
    let threw = false;
    try {
        evalVBASingle(code);
    } catch (_e) {
        threw = true;
    }
    assert.ok(threw, `[FAIL] ${label}: コンパイルエラーが発生しなかった`);
    console.log(`[PASS] ${label}`);
}

function expectNoError(code: string, label: string): void {
    try {
        evalVBASingle(code);
    } catch (e: any) {
        assert.fail(`[FAIL] ${label}: 予期しないエラー: ${e.message}`);
    }
    console.log(`[PASS] ${label}`);
}

// ---------------------------------------------------------------------------
// For...Next: 変数不一致はコンパイルエラー
// ---------------------------------------------------------------------------
expectCompileError(`
Sub Test()
    Dim i As Long, j As Long
    For i = 0 To 10
    Next j
End Sub`, 'For i...Next j → コンパイルエラー');

expectCompileError(`
Sub Test()
    Dim i As Long, j As Long
    For i = 0 To 10 : Next j
End Sub`, 'For i...Next j（1行形式）→ コンパイルエラー');

// ---------------------------------------------------------------------------
// For...Next: 変数一致・Next のみは正常
// ---------------------------------------------------------------------------
expectNoError(`
Sub Test()
    Dim i As Long
    For i = 0 To 10
    Next i
End Sub`, 'For i...Next i → 正常');

expectNoError(`
Sub Test()
    Dim i As Long
    For i = 0 To 10
    Next
End Sub`, 'For i...Next（変数なし）→ 正常');

// 大文字小文字を区別しない
expectNoError(`
Sub Test()
    Dim i As Long
    For i = 0 To 10
    Next I
End Sub`, 'For i...Next I（大文字）→ 正常');

// ---------------------------------------------------------------------------
// For Each...Next: 変数不一致はコンパイルエラー
// ---------------------------------------------------------------------------
expectCompileError(`
Sub Test()
    Dim arr(3) As Long
    Dim x As Variant, y As Variant
    For Each x In arr
    Next y
End Sub`, 'For Each x...Next y → コンパイルエラー');

// ---------------------------------------------------------------------------
// For Each...Next: 変数一致・Next のみは正常
// ---------------------------------------------------------------------------
expectNoError(`
Sub Test()
    Dim arr(3) As Long
    Dim x As Variant
    For Each x In arr
    Next x
End Sub`, 'For Each x...Next x → 正常');

expectNoError(`
Sub Test()
    Dim arr(3) As Long
    Dim x As Variant
    For Each x In arr
    Next
End Sub`, 'For Each x...Next（変数なし）→ 正常');

// ---------------------------------------------------------------------------
// ネスト: 内側・外側それぞれ正しい変数なら正常
// ---------------------------------------------------------------------------
expectNoError(`
Sub Test()
    Dim i As Long, j As Long
    For i = 0 To 3
        For j = 0 To 3
        Next j
    Next i
End Sub`, 'ネスト For i/j...Next j/i → 正常');

// ---------------------------------------------------------------------------
// ネスト: 内側の Next が外側の変数 → コンパイルエラー
// ---------------------------------------------------------------------------
expectCompileError(`
Sub Test()
    Dim i As Long, j As Long
    For i = 0 To 3
        For j = 0 To 3
        Next i
    Next i
End Sub`, 'ネスト: 内側の Next i（外側変数）→ コンパイルエラー');

console.log('\n✅ for-next-variable-mismatch: 全テスト通過');
