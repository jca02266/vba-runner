import { evalVBASingle, assert } from '../../test-libs/test-runner';

// CVErr and IsError tests
{
    const code = `
    Sub Test()
        Dim v As Variant
        v = CVErr(11) ' Division by zero
        Debug.Print IsError(v)
        Debug.Print Typename(v)
        
        If v = CVErr(11) Then
            Debug.Print "Matches 11"
        End If
        
        ' Range check
        On Error Resume Next
        CVErr(70000)
        If Err.Number <> 0 Then Debug.Print "Error:" & Err.Number
    End Sub
    `;
    let output = "";
    const ev = evalVBASingle(code, { onPrint: s => output += s + "\n" });
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "True", "IsError(v) should be True");
    assert.strictEqual(lines[1], "Error", "Typename(v) should be 'Error'");
    assert.strictEqual(lines[2], "Matches 11", "Comparison of CVErr values should work");
    assert.strictEqual(lines[3], "Error:5", "CVErr(70000) should raise Error 5 (Invalid procedure call or argument)");
    console.log('[PASS] CVErr and IsError tests');
}

console.log('\n✅ CVErr: 全テスト通過');
