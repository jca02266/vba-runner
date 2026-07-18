import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: New Collection and Count
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add 10
    col.Add 20
    col.Add 30
    F = col.Count
End Function
`, 'F');
    assert.strictEqual(result, 3, 'Count after 3 adds');
    console.log('[PASS] New Collection: Count = 3');
}

// Test 2: Item by 1-based index
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "apple"
    col.Add "banana"
    col.Add "cherry"
    F = col.Item(2)
End Function
`, 'F');
    assert.strictEqual(result, 'banana', 'Item(2) returns second element');
    console.log('[PASS] Item(index): 1-based access');
}

// Test 3: Add with Key, Item by string key
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "Tokyo", "TYO"
    col.Add "Osaka", "OSA"
    F = col.Item("TYO")
End Function
`, 'F');
    assert.strictEqual(result, 'Tokyo', 'Item by string key');
    console.log('[PASS] Add with Key: Item by string key');
}

// Test 4: Remove by index
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "a"
    col.Add "b"
    col.Add "c"
    col.Remove 2
    F = col.Count & "," & col.Item(2)
End Function
`, 'F');
    assert.strictEqual(result, '2,c', 'Remove(2) shifts remaining items');
    console.log('[PASS] Remove(index): item removed, Count decremented');
}

// Test 5: Remove by key
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "x", "alpha"
    col.Add "y", "beta"
    col.Remove "alpha"
    F = col.Count
End Function
`, 'F');
    assert.strictEqual(result, 1, 'Remove by key');
    console.log('[PASS] Remove(key): item removed by string key');
}

// Test 6: Add with Before
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "a"
    col.Add "c"
    col.Add "b", , 2
    F = col.Item(1) & col.Item(2) & col.Item(3)
End Function
`, 'F');
    assert.strictEqual(result, 'abc', 'Add Before=2 inserts before second element');
    console.log('[PASS] Add Before: inserted at correct position');
}

// Test 7: Add with After
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    col.Add "a"
    col.Add "c"
    col.Add "b", , , 1
    F = col.Item(1) & col.Item(2) & col.Item(3)
End Function
`, 'F');
    assert.strictEqual(result, 'abc', 'Add After=1 inserts after first element');
    console.log('[PASS] Add After: inserted at correct position');
}

// Test 8: For Each iteration
{
    const result = runFunc(`
Function F()
    Dim col As New Collection
    Dim x As Variant
    Dim s As String
    col.Add 1
    col.Add 2
    col.Add 3
    For Each x In col
        s = s & x
    Next x
    F = s
End Function
`, 'F');
    assert.strictEqual(result, '123', 'For Each iterates over Collection items');
    console.log('[PASS] For Each: iterates all items in order');
}

// Test 9: Duplicate key raises error 457
{
    let caught = false;
    try {
        runFunc(`
Sub S()
    Dim col As New Collection
    col.Add "a", "key1"
    col.Add "b", "key1"
End Sub
`, 'S');
    } catch (e: any) {
        caught = true;
        assert.strictEqual(e.number, 457, 'Duplicate key raises error 457');
    }
    assert.strictEqual(caught, true, 'Error was raised');
    console.log('[PASS] Duplicate key: raises error 457');
}

// Test 10: Set col = New Collection (explicit Set)
{
    const result = runFunc(`
Function F()
    Dim col As Collection
    Set col = New Collection
    col.Add 99
    F = col.Item(1)
End Function
`, 'F');
    assert.strictEqual(result, 99, 'Set col = New Collection works');
    console.log('[PASS] Set col = New Collection: explicit Set works');
}

// Bug 33-B（評価 #33）: Collection.Add の名前付き Before:=/After:= が黙って無視され、
// 名前付き引数の値が位置引数（Key）に化けていた
{
    const code = `
Function TestNamedInsert() As String
    Dim c As New Collection
    c.Add "A"
    c.Add "B"
    c.Add "C", After:=1
    c.Add "Z", Before:=1
    Dim s As String, v As Variant
    For Each v In c
        s = s & v
    Next
    TestNamedInsert = s
End Function
`;
    assert.strictEqual(runFunc(code, 'TestNamedInsert'), 'ZACB',
        'Bug 33-B: Collection.Add の Before:=/After:= 名前付き引数で挿入位置が効く');
    console.log('[PASS] Bug 33-B: Collection.Add の名前付き Before/After');
}

// Bug 33-B 続: パラメーター仕様のない関数への名前付き引数は Error 448 で明示的に失敗する
{
    const code = `
Sub TestUnknownNamed()
    Dim c As New Collection
    c.Add "A", Bogus:=1
End Sub
`;
    assert.throwsMatch(() => runFunc(code, 'TestUnknownNamed'), /error '448'/,
        'Bug 33-B: 未知の名前付き引数は Error 448');
    console.log('[PASS] Bug 33-B: 未知の名前付き引数は Error 448');
}

console.log('\n✅ collection: 全テスト通過');
