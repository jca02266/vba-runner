import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Set dict.Item(key) = obj（バグ再現: evaluateSetStatement が CallExpression LHS を拒否していた）
{
    const result = runFunc(`
Function F() As String
    Dim outer As Object
    Set outer = CreateObject("Scripting.Dictionary")
    Dim inner As Object
    Set inner = CreateObject("Scripting.Dictionary")
    inner("val") = "hello"
    Set outer.Item("sub") = inner
    F = outer("sub")("val")
End Function
`, 'F');
    assert.strictEqual(result, 'hello', 'Set dict.Item(key) = obj: nested Dictionary via Set member call');
    console.log('[PASS] Set dict.Item(key) = obj');
}

// Test 2: Set dict(key) = obj（Identifier callee の Set 代入）
{
    const result = runFunc(`
Function F() As Long
    Dim outer As Object
    Set outer = CreateObject("Scripting.Dictionary")
    Dim inner As Object
    Set inner = CreateObject("Scripting.Dictionary")
    inner("n") = 99
    Set outer("sub") = inner
    F = outer("sub")("n")
End Function
`, 'F');
    assert.strictEqual(result, 99, 'Set dict(key) = obj: nested Dictionary via Set call');
    console.log('[PASS] Set dict(key) = obj');
}

// Test 3: outer("sub")("x") = val（バグ再現: callee が CallExpression のチェーン代入）
{
    const result = runFunc(`
Function F() As Long
    Dim outer As Object
    Set outer = CreateObject("Scripting.Dictionary")
    Dim inner As Object
    Set inner = CreateObject("Scripting.Dictionary")
    Set outer("sub") = inner
    outer("sub")("x") = 42
    F = outer("sub")("x")
End Function
`, 'F');
    assert.strictEqual(result, 42, 'outer(key1)(key2) = val: chained dictionary assignment');
    console.log('[PASS] outer("sub")("x") = val chain');
}

// Test 4: 3段ネスト dict(k1)(k2)(k3) = val
{
    const result = runFunc(`
Function F() As String
    Dim d1 As Object
    Set d1 = CreateObject("Scripting.Dictionary")
    Dim d2 As Object
    Set d2 = CreateObject("Scripting.Dictionary")
    Dim d3 As Object
    Set d3 = CreateObject("Scripting.Dictionary")
    d3("leaf") = "deep"
    Set d2("c") = d3
    Set d1("b") = d2
    F = d1("b")("c")("leaf")
End Function
`, 'F');
    assert.strictEqual(result, 'deep', '3-level nested dict read');
    console.log('[PASS] 3-level nested dict read');
}

// Test 5: For Each key In dict.Keys でキー列挙
{
    const result = runFunc(`
Function F() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("a") = 1
    d("b") = 2
    d("c") = 3
    Dim k As Variant
    Dim s As String
    For Each k In d.Keys
        s = s & k
    Next k
    F = s
End Function
`, 'F');
    assert.strictEqual(result, 'abc', 'For Each key In dict.Keys enumerates all keys');
    console.log('[PASS] For Each key In dict.Keys');
}

console.log('\n✅ object-member-assignment: 全テスト通過');
