/**
 * For Each Statement のテスト (§5.4.2.4)
 *
 * 網羅する観点:
 *   1. 1次元配列の反復
 *   2. Collection オブジェクトの反復
 *   3. Dictionary オブジェクトのキー反復
 *   4. 空コレクションでのスキップ
 *   5. Exit For による早期脱出
 *   6. Next に変数名を付けた明示的形式
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. 1次元配列の反復 ---
const arrayCode = `
Function SumArray()
    Dim arr(4)
    arr(0) = 10
    arr(1) = 20
    arr(2) = 30
    arr(3) = 40
    arr(4) = 50
    Dim total
    total = 0
    Dim x
    For Each x In arr
        total = total + x
    Next
    SumArray = total
End Function
`;
assert.strictEqual(runFunc(arrayCode, 'SumArray'), 150, '1次元配列の合計');
console.log('[PASS] 1次元配列の反復');

// --- 2. Collection オブジェクトの反復 ---
const colCode = `
Function SumCollection()
    Dim col As New Collection
    col.Add 10
    col.Add 20
    col.Add 30
    Dim total
    total = 0
    Dim x
    For Each x In col
        total = total + x
    Next
    SumCollection = total
End Function
`;
assert.strictEqual(runFunc(colCode, 'SumCollection'), 60, 'Collection の合計');
console.log('[PASS] Collection オブジェクトの反復');

// --- 3. Dictionary キーの反復 ---
const dictCode = `
Function ConcatKeys()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d("apple") = 1
    d("banana") = 2
    d("cherry") = 3
    Dim result
    result = ""
    Dim k
    For Each k In d
        result = result & k & ","
    Next
    ConcatKeys = result
End Function
`;
const keysResult = runFunc(dictCode, 'ConcatKeys');
assert.strictEqual(typeof keysResult, 'string', 'Dictionary キー反復: 文字列を返す');
assert.strictEqual(keysResult.includes('apple'), true, 'apple キーが含まれる');
assert.strictEqual(keysResult.includes('banana'), true, 'banana キーが含まれる');
assert.strictEqual(keysResult.includes('cherry'), true, 'cherry キーが含まれる');
console.log('[PASS] Dictionary キーの反復');

// --- 4. 空 Collection でのスキップ ---
const emptyCode = `
Function TestEmpty()
    Dim col As New Collection
    Dim count
    count = 0
    Dim x
    For Each x In col
        count = count + 1
    Next
    TestEmpty = count
End Function
`;
assert.strictEqual(runFunc(emptyCode, 'TestEmpty'), 0, '空 Collection は反復しない');
console.log('[PASS] 空コレクションでのスキップ');

// --- 5. Exit For による早期脱出 ---
const exitForCode = `
Function FirstNegative()
    Dim arr(4)
    arr(0) = 5
    arr(1) = 3
    arr(2) = -1
    arr(3) = 7
    arr(4) = -2
    Dim result
    result = 0
    Dim x
    For Each x In arr
        If x < 0 Then
            result = x
            Exit For
        End If
    Next
    FirstNegative = result
End Function
`;
assert.strictEqual(runFunc(exitForCode, 'FirstNegative'), -1, 'Exit For で最初の負の値で停止');
console.log('[PASS] Exit For による早期脱出');

// --- 6. 明示的 Next 変数名 ---
const explicitNextCode = `
Function SumExplicit()
    Dim arr(2)
    arr(0) = 1
    arr(1) = 2
    arr(2) = 3
    Dim total
    total = 0
    Dim item
    For Each item In arr
        total = total + item
    Next item
    SumExplicit = total
End Function
`;
assert.strictEqual(runFunc(explicitNextCode, 'SumExplicit'), 6, 'Next に変数名を付けても正常動作');
console.log('[PASS] 明示的 Next 変数名');

// --- 7. Collection のループ内で要素の文字列化 ---
const strColCode = `
Function JoinItems()
    Dim col As New Collection
    col.Add "foo"
    col.Add "bar"
    col.Add "baz"
    Dim result
    result = ""
    Dim s
    For Each s In col
        If result <> "" Then result = result & ","
        result = result & s
    Next
    JoinItems = result
End Function
`;
assert.strictEqual(runFunc(strColCode, 'JoinItems'), 'foo,bar,baz', 'Collection 文字列要素の連結');
console.log('[PASS] Collection 文字列要素の連結');

// --- 8. Set obj = New Collection から返した Collection の For Each ---
// New Collection を Set で受け渡す場合、instantiateClass 経由の外部ファクトリオブジェクトになる
// items プロパティではなく Symbol.iterator を使って列挙できること
const colFromFuncCode = `
Function MakeCol() As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add "p"
    c.Add "q"
    c.Add "r"
    Set MakeCol = c
End Function

Function F() As String
    Dim acc As String
    Dim v As Variant
    For Each v In MakeCol()
        acc = acc & v
    Next v
    F = acc
End Function
`;
assert.strictEqual(runFunc(colFromFuncCode, 'F'), 'pqr', 'Set = New Collection から返した Collection の For Each');
console.log('[PASS] Set = New Collection から返した Collection の For Each');

console.log('\n✅ For Each Statement: 全テスト通過');
