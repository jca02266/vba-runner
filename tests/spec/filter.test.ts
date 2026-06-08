import { vbaTrue, vbaFalse } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// Basic Filter tests
{
    const code = `
    Sub Test()
        Dim arr() As Variant
        arr = Array("apple", "banana", "cherry", "date", "apricot")
        
        Dim res() As Variant
        res = Filter(arr, "ap")
        Debug.Print "Count1:" & UBound(res) + 1
        Debug.Print "Item1:" & res(0)
        Debug.Print "Item2:" & res(1)
        
        Dim res2() As Variant
        res2 = Filter(arr, "ap", False)
        Debug.Print "Count2:" & UBound(res2) + 1
        Debug.Print "Item2-1:" & res2(0)
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "Count1:2", "Should find 'apple' and 'apricot'");
    assert.strictEqual(lines[1], "Item1:apple");
    assert.strictEqual(lines[2], "Item2:apricot");
    
    assert.strictEqual(lines[3], "Count2:3", "Should find 'banana', 'cherry', 'date'");
    assert.strictEqual(lines[4], "Item2-1:banana");
    console.log('[PASS] Basic Filter tests');
}

// Case sensitivity tests
{
    const code = `
    Sub Test()
        Dim arr() As Variant
        arr = Array("Apple", "banana", "APPLE")
        
        ' Binary compare (default or vbBinaryCompare=0)
        Dim res1() As Variant
        res1 = Filter(arr, "Apple")
        Debug.Print "BinaryCount:" & UBound(res1) + 1
        
        ' Text compare (vbTextCompare=1)
        Dim res2() As Variant
        res2 = Filter(arr, "Apple", True, 1)
        Debug.Print "TextCount:" & UBound(res2) + 1
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "BinaryCount:1", "Binary should match exactly 'Apple'");
    assert.strictEqual(lines[1], "TextCount:2", "Text should match 'Apple' and 'APPLE'");
    console.log('[PASS] Case sensitivity tests');
}

// Edge cases
{
    const code = `
    Sub Test()
        Dim arr() As Variant
        arr = Array() ' Empty array
        
        Dim res1() As Variant
        res1 = Filter(arr, "a")
        Debug.Print "EmptyCount:" & UBound(res1) + 1
        
        arr = Array("a", "b")
        Dim res2() As Variant
        res2 = Filter(arr, "z")
        Debug.Print "NoMatchCount:" & UBound(res2) + 1
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "EmptyCount:0");
    assert.strictEqual(lines[1], "NoMatchCount:0");
    console.log('[PASS] Edge case tests');
}

console.log('\n✅ Filter: 全テスト通過');
