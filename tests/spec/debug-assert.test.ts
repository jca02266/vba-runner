import { evalVBASingle, assert as specAssert } from '../../test-libs/test-runner';

// Debug.Assert tests
{
    const code = `
    Sub Test()
        Debug.Assert True
        Debug.Print "Pass1"
        ' Debug.Assert False ' This should throw or stop
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    specAssert.strictEqual(lines[0], "Pass1", "Debug.Assert True should pass");
    console.log('[PASS] Debug.Assert True');
}

{
    const code = `
    Sub Test()
        On Error Resume Next
        Debug.Assert False
        If Err.Number <> 0 Then Debug.Print "Error:" & Err.Description
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    specAssert.strictEqual(lines[0], "Error:Assertion failed", "Debug.Assert False should raise an error");
    console.log('[PASS] Debug.Assert False');
}

// --- Bug BR: Debug.Print で `;` セパレーターが ParseError になっていた ---
{
    const outputs: string[] = [];
    const ev = evalVBASingle(`
        Sub Test()
            Debug.Print "Hello"; "World"
            Debug.Print "A", "B"
            Debug.Print 1; 2; 3
            Debug.Print Spc(3); "Y"
        End Sub
    `, { onPrint: (s: string) => outputs.push(s) });
    ev.callProcedure('Test', []);
    specAssert.strictEqual(outputs[0], 'HelloWorld', 'Debug.Print "; " セミコロンは連結');
    specAssert.ok(outputs[1].startsWith('A') && outputs[1].includes('B'), 'Debug.Print "," コンマはタブ区切り');
    specAssert.strictEqual(outputs[2], '123', 'Debug.Print 数値 "; "');
    specAssert.strictEqual(outputs[3], '   Y', 'Debug.Print Spc(3); "Y"');
    console.log('[PASS] Bug BR: Debug.Print "; " セミコロンセパレーター');
}

console.log('\n✅ Debug.Assert: 全テスト通過');
