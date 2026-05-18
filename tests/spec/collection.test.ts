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

console.log('\n✅ collection: 全テスト通過');
