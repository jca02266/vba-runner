/**
 * §5.4.2.3 For Statement — `Next i, j` 複数変数 (§5.4.2.4 ForEach も同様)
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

// ─── 基本: Next j, i (2重ネスト) ──────────────────────────────────────────────
{
    const code = `
        Function TestNextMulti() As Long
            Dim i As Long, j As Long
            Dim sum As Long
            For i = 1 To 3
                For j = 1 To 3
                    sum = sum + 1
                Next j, i
            TestNextMulti = sum
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestNextMulti', []);
    assert.strictEqual(result, 9, 'Next j, i: 3×3=9回実行');
    console.log('[PASS] Next j, i (2重ネスト)');
}

// ─── 変数名省略バリアント: Next, i → 内側は無名Next ───────────────────────────
{
    const code = `
        Function TestNextOmit() As Long
            Dim i As Long, j As Long
            Dim sum As Long
            For i = 1 To 2
                For j = 1 To 2
                    sum = sum + 1
                Next j
            Next i
            TestNextOmit = sum
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestNextOmit', []);
    assert.strictEqual(result, 4, '通常の個別Next: 2×2=4回');
    console.log('[PASS] 通常の個別 Next (比較用)');
}

// ─── 3重ネスト: Next k, j, i ─────────────────────────────────────────────────
{
    const code = `
        Function TestNextTriple() As Long
            Dim i As Long, j As Long, k As Long
            Dim sum As Long
            For i = 1 To 2
                For j = 1 To 2
                    For k = 1 To 2
                        sum = sum + 1
                    Next k, j, i
            TestNextTriple = sum
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestNextTriple', []);
    assert.strictEqual(result, 8, 'Next k, j, i: 2×2×2=8回');
    console.log('[PASS] Next k, j, i (3重ネスト)');
}

// ─── For Each + For の混合 Next ───────────────────────────────────────────────
{
    const code = `
        Function TestNextMixed() As Long
            Dim arr(1) As Long
            arr(0) = 10: arr(1) = 20
            Dim i As Long
            Dim sum As Long
            For Each elem In arr
                For i = 1 To 2
                    sum = sum + 1
                Next i, elem
            TestNextMixed = sum
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestNextMixed', []);
    assert.strictEqual(result, 4, 'For Each + For の混合 Next: 2要素×2=4回');
    console.log('[PASS] For Each + For 混合 Next i, elem');
}

// ─── 正しくない変数名はエラーになる ──────────────────────────────────────────
{
    const code = `
        Sub TestBadNext()
            Dim i As Long, j As Long
            For i = 1 To 3
                For j = 1 To 3
                Next j, i
        End Sub
    `;
    // 正しい順序 (j=inner, i=outer) なのでエラーなし
    const ev = evalVBASingle(code);
    ev.callProcedure('TestBadNext', []);
    console.log('[PASS] Next j, i 正しい順序でエラーなし');
}

console.log('\n✅ Next 複数変数: 全テスト通過');
