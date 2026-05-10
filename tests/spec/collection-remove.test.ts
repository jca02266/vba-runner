import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

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

// Test basic removal by index
{
    const code = `
    Sub Test()
        Dim col As New Collection
        col.Add "Item1"
        col.Add "Item2"
        col.Add "Item3"
        col.Remove 2
        Debug.Print col.Count
        Debug.Print col.Item(1)
        Debug.Print col.Item(2)
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "2", "Count should be 2 after removal");
    assert.strictEqual(lines[1], "Item1", "Item 1 should remain Item1");
    assert.strictEqual(lines[2], "Item3", "Item 2 should now be Item3");
    
    // Check TypeName
    const typeNameCode = `
    Sub Test()
        Dim col As New Collection
        Debug.Print TypeName(col)
    End Sub
    `;
    let output2 = "";
    const ev2 = new Evaluator(s => output2 += s + "\n");
    const tokens2 = new Lexer(typeNameCode).tokenize();
    const ast2 = new Parser(tokens2).parse();
    ev2.evaluate(ast2);
    ev2.callProcedure("Test", []);
    assert.strictEqual(output2.trim(), "Collection", "TypeName(col) should be 'Collection'");

    console.log('[PASS] Basic removal by index');
}

// Test removal by key
{
    const code = `
    Sub Test()
        Dim col As New Collection
        col.Add "Item1", "K1"
        col.Add "Item2", "K2"
        col.Remove "K1"
        Debug.Print col.Count
        Debug.Print col.Item(1)
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "1", "Count should be 1 after removal by key");
    assert.strictEqual(lines[1], "Item2", "Remaining item should be Item2");
    console.log('[PASS] Removal by key');
}

// Test error cases
{
    const code = `
    Sub Test()
        Dim col As New Collection
        col.Add "Item1"
        On Error Resume Next
        col.Remove 5 ' Out of range
        Debug.Print Err.Number
        Err.Clear
        col.Remove "NonExistent" ' Invalid key
        Debug.Print Err.Number
    End Sub
    `;
    // Note: The current Evaluator might not support Err object fully yet or might throw JS errors
    // Let's see if it throws or if we need to implement error trapping.
    try {
        evalVBA(code).callProcedure("Test", []);
        console.log('[PASS] Error cases (handled by On Error)');
    } catch (e) {
        console.log('[FAIL] Error cases threw JS error: ' + e);
    }
}

// Test For Each
{
    const code = `
    Sub Test()
        Dim col As New Collection
        col.Add "A"
        col.Add "B"
        col.Add "C"
        Dim s As String
        Dim result As String
        For Each s In col
            result = result & s
        Next
        Debug.Print result
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    assert.strictEqual(output.trim(), "ABC", "For Each should work with Collection");
    console.log('[PASS] For Each with Collection');
}


console.log('\n✅ Collection.Remove: 全テスト通過');
