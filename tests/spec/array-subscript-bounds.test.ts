/**
 * 配列の添字境界チェック (Error 9 Subscript out of range)
 *
 * - 1D / 2D 配列の読み取り・書き込み双方で境界を超えたとき Error 9 を投げる
 * - ReDim 後も正しい境界を使う
 * - Option Base 1 で lower=1 になるため index 0 は Error 9
 * - 次元数不一致も Error 9
 * - 境界内アクセスはエラーなし
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function catchError(code: string, entry: string): any {
    const ev = evalVBASingle(code);
    try {
        ev.callProcedure(entry, []);
    } catch (e: any) {
        return e;
    }
    assert.fail(`expected ${entry} to throw`);
    throw new Error('unreachable');
}

function noError(code: string, entry: string): void {
    const ev = evalVBASingle(code);
    try {
        ev.callProcedure(entry, []);
    } catch (e: any) {
        assert.fail(`expected ${entry} not to throw, but got Error ${e.number} line ${e.vbaLine}`);
    }
}

// ---------------------------------------------------------------------------
// 1D 読み取り
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 5) As Long
    Dim x As Long
    x = arr(6)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '1D read over upper: Error 9');
    assert.strictEqual(e.vbaLine, 5, '1D read over upper: line 5');
    console.log('[PASS] 1D 読み取り: 上限超え');
}
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 5) As Long
    Dim x As Long
    x = arr(0)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '1D read under lower: Error 9');
    assert.strictEqual(e.vbaLine, 5, '1D read under lower: line 5');
    console.log('[PASS] 1D 読み取り: 下限未満');
}

// ---------------------------------------------------------------------------
// 1D 書き込み
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 5) As Long
    arr(6) = 99
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '1D write over upper: Error 9');
    assert.strictEqual(e.vbaLine, 4, '1D write over upper: line 4');
    console.log('[PASS] 1D 書き込み: 上限超え');
}
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 5) As Long
    arr(0) = 99
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '1D write under lower: Error 9');
    assert.strictEqual(e.vbaLine, 4, '1D write under lower: line 4');
    console.log('[PASS] 1D 書き込み: 下限未満');
}

// ---------------------------------------------------------------------------
// 1D 境界内は正常
// ---------------------------------------------------------------------------
{
    noError(`
Sub Test()
    Dim arr(1 To 5) As Long
    arr(1) = 10
    arr(5) = 50
    Dim x As Long
    x = arr(1)
    x = arr(5)
End Sub`, 'Test');
    console.log('[PASS] 1D 境界内アクセス: エラーなし');
}

// ---------------------------------------------------------------------------
// 2D 読み取り
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    Dim x As Long
    x = arr(4, 1)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '2D read dim1 over: Error 9');
    console.log('[PASS] 2D 読み取り: 第1次元 上限超え');
}
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    Dim x As Long
    x = arr(1, 5)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '2D read dim2 over: Error 9');
    console.log('[PASS] 2D 読み取り: 第2次元 上限超え');
}

// ---------------------------------------------------------------------------
// 2D 書き込み
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    arr(4, 1) = 99
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '2D write dim1 over: Error 9');
    console.log('[PASS] 2D 書き込み: 第1次元 上限超え');
}
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    arr(1, 5) = 99
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '2D write dim2 over: Error 9');
    console.log('[PASS] 2D 書き込み: 第2次元 上限超え');
}

// ---------------------------------------------------------------------------
// 2D 境界内は正常
// ---------------------------------------------------------------------------
{
    noError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    arr(1, 1) = 1
    arr(3, 4) = 99
    Dim x As Long
    x = arr(3, 4)
End Sub`, 'Test');
    console.log('[PASS] 2D 境界内アクセス: エラーなし');
}

// ---------------------------------------------------------------------------
// 次元数不一致
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr(1 To 3, 1 To 4) As Long
    Dim x As Long
    x = arr(1)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, '2D accessed with 1 index: Error 9');
    console.log('[PASS] 2D 配列を 1 インデックスでアクセス: Error 9');
}

// ---------------------------------------------------------------------------
// ReDim 後の境界チェック
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Sub Test()
    Dim arr() As Long
    ReDim arr(1 To 3)
    arr(3) = 99
    arr(4) = 1
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, 'ReDim bounds: Error 9');
    assert.strictEqual(e.vbaLine, 6, 'ReDim bounds: line 6');
    console.log('[PASS] ReDim 後の境界チェック');
}
{
    noError(`
Sub Test()
    Dim arr() As Long
    ReDim arr(1 To 3)
    arr(1) = 10
    arr(3) = 30
End Sub`, 'Test');
    console.log('[PASS] ReDim 後 境界内アクセス: エラーなし');
}

// ---------------------------------------------------------------------------
// Option Base 1 — index 0 は Error 9
// ---------------------------------------------------------------------------
{
    const e = catchError(`
Option Base 1
Sub Test()
    Dim arr(3) As Long
    arr(0) = 1
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, 'Base1 zero write: Error 9');
    console.log('[PASS] Option Base 1: index 0 書き込みは Error 9');
}
{
    const e = catchError(`
Option Base 1
Sub Test()
    Dim arr(3) As Long
    Dim x As Long
    x = arr(0)
End Sub`, 'Test');
    assert.strictEqual(e.number, 9, 'Base1 zero read: Error 9');
    console.log('[PASS] Option Base 1: index 0 読み取りは Error 9');
}
{
    noError(`
Option Base 1
Sub Test()
    Dim arr(3) As Long
    arr(1) = 10
    arr(3) = 30
End Sub`, 'Test');
    console.log('[PASS] Option Base 1: 境界内アクセス エラーなし');
}

console.log('\n✅ array-subscript-bounds: 全テスト通過');
