/**
 * RangeAccessDetector (src/engine/range-access-detector.ts) のテスト
 * Phase 3: Range 変数経由のアクセス検出
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { detectRangeAccess } from '../../src/engine/range-access-detector';
import { assert } from '../../test-libs/test-runner';

function detect(code: string) {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    return detectRangeAccess(ast);
}

// ─── 1. member-call: rng.Item(r, c) ──────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            rng.Item(1, 2)
        End Sub
    `);

    assert.strictEqual(hits.length, 1, 'hits = 1');
    assert.strictEqual(hits[0].kind, 'member-call', 'kind = member-call');
    assert.strictEqual(hits[0].varName, 'rng', 'varName = rng');
    assert.strictEqual(hits[0].property, 'Item', 'property = Item');
    console.log('[PASS] member-call: rng.Item()');
}

// ─── 2. index-call: rng(r, c) ────────────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            rng(1, 2)
        End Sub
    `);

    assert.strictEqual(hits.length, 1, 'hits = 1');
    assert.strictEqual(hits[0].kind, 'index-call', 'kind = index-call');
    assert.strictEqual(hits[0].varName, 'rng', 'varName = rng');
    assert.strictEqual(hits[0].property, undefined, 'property = undefined');
    console.log('[PASS] index-call: rng(r,c)');
}

// ─── 3. member-access: rng.Value ─────────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            rng.Value = 10
        End Sub
    `);

    const memberAccess = hits.filter(h => h.kind === 'member-access');
    assert.strictEqual(memberAccess.length >= 1, true, 'member-access ヒットあり');
    assert.strictEqual(memberAccess[0].property, 'Value', 'property = Value');
    console.log('[PASS] member-access: rng.Value');
}

// ─── 4. 非 Range 変数は検出されない ─────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim n As Long
            n = n + 1
            Dim s As String
            s = "hello"
        End Sub
    `);

    assert.strictEqual(hits.length, 0, '非 Range 変数はヒットしない');
    console.log('[PASS] 非 Range 変数は検出されない');
}

// ─── 5. フロー追跡: Set で Range になった変数を検出 ──────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim r As Variant
            Set r = ActiveSheet.Cells
            r(1, 1) = 99
        End Sub
    `);

    const indexHits = hits.filter(h => h.kind === 'index-call' && h.varName === 'r');
    assert.strictEqual(indexHits.length, 1, 'フロー追跡で Range 検出');
    console.log('[PASS] フロー追跡: Set で Range 変数を検出');
}

// ─── 6. 複数ヒット ────────────────────────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            rng.Item(1, 1)
            rng.Value = 10
            rng(2, 2)
        End Sub
    `);

    assert.strictEqual(hits.length >= 3, true, '3 件以上のヒット');
    const kinds = hits.map(h => h.kind);
    assert.strictEqual(kinds.includes('member-call'),  true, 'member-call あり');
    assert.strictEqual(kinds.includes('member-access'), true, 'member-access あり');
    assert.strictEqual(kinds.includes('index-call'),   true, 'index-call あり');
    console.log('[PASS] 複数ヒット');
}

// ─── 7. パラメーターが Range 型の場合も検出 ──────────────────────────────────
{
    const hits = detect(`
        Sub Process(rng As Range)
            rng.Value = 0
        End Sub
    `);

    assert.strictEqual(hits.length >= 1, true, 'パラメーター Range のヒット');
    assert.strictEqual(hits[0].varName, 'rng', 'varName = rng');
    console.log('[PASS] パラメーター Range の検出');
}

// ─── 8. If / For 内のアクセスも検出 ─────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            If True Then
                rng.Item(1, 1)
            End If
            For i = 1 To 10
                rng(i, 1) = i
            Next i
        End Sub
    `);

    const memberHits = hits.filter(h => h.kind === 'member-call');
    const indexHits  = hits.filter(h => h.kind === 'index-call');
    assert.strictEqual(memberHits.length >= 1, true, 'If 内の member-call 検出');
    assert.strictEqual(indexHits.length >= 1,  true, 'For 内の index-call 検出');
    console.log('[PASS] If/For 内のアクセス検出');
}

// ─── 9. 行番号が記録されている ────────────────────────────────────────────────
{
    const hits = detect(`
        Sub Test()
            Dim rng As Range
            rng.Item(1, 1)
        End Sub
    `);

    assert.strictEqual(typeof hits[0].line,   'number', 'line は number');
    assert.strictEqual(typeof hits[0].column, 'number', 'column は number');
    console.log('[PASS] 行番号・列番号の記録');
}

// ─── 10. モジュールレベル Range 変数のアクセス検出 ───────────────────────────
{
    const hits = detect(`
        Dim gRng As Range
        Sub Test()
            gRng.Item(1, 1)
        End Sub
    `);

    const moduleHits = hits.filter(h => h.varName === 'gRng');
    assert.strictEqual(moduleHits.length >= 1, true, 'モジュールレベル変数の検出');
    console.log('[PASS] モジュールレベル Range 変数の検出');
}

console.log('\n✅ RangeAccessDetector: 全テスト通過');
