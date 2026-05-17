import { Evaluator } from '../../../src/engine/evaluator';
import { Parser } from '../../../src/engine/parser';
import { Lexer } from '../../../src/engine/lexer';

function assert(condition: boolean, message: string) {
    if (!condition) {
        console.error(`[FAIL] ${message}`);
        process.exit(1);
    } else {
        console.log(`[PASS] ${message}`);
    }
}

async function testErrorHandling() {
    console.log("--- Testing Error Handling ---");
    
    const vbaCode = `
        Sub TestError()
            Dim x As Integer
            On Error GoTo ErrorHandler
            x = 1 / 0
            Debug.Print "This should not be printed"
            Exit Sub
        ErrorHandler:
            Debug.Print "Caught error: " & Err.Number
            Resume Next
            Debug.Print "This should not be printed either"
        End Sub
        
        Sub TestResumeNext()
            On Error Resume Next
            Dim x As Integer
            x = 1 / 0
            Debug.Print "After error: " & Err.Number
        End Sub

        Sub TestGoTo0()
            On Error Resume Next
            On Error GoTo 0
            Dim x As Integer
            x = 1 / 0
            Debug.Print "This should not be reached"
        End Sub
    `;

    const logs: string[] = [];
    const evaluator = new Evaluator((msg) => logs.push(msg));
    
    const lexer = new Lexer(vbaCode);
    const parser = new Parser(lexer.tokenize());
    const program = parser.parse();
    
    evaluator.evaluate(program);
    
    console.log("Running TestError...");
    evaluator.callProcedure('TestError', []);
    assert(logs.includes("Caught error: 11"), "ErrorHandler should catch division by zero (Error 11)");
    assert(logs.includes("This should not be printed"), "Resume Next should continue at next line");

    console.log("Running TestResumeNext...");
    evaluator.callProcedure('TestResumeNext', []);
    assert(logs.includes("After error: 11"), "TestResumeNext should work and print error number");

    console.log("Running TestGoTo0...");
    try {
        evaluator.callProcedure('TestGoTo0', []);
        assert(false, "TestGoTo0 should have thrown an error");
    } catch (e) {
        assert(true, "TestGoTo0 correctly threw an error after On Error GoTo 0");
    }

    console.log("--- All Error Handling Tests Passed! ---");
}

testErrorHandling().catch(err => {
    console.error(err);
    process.exit(1);
});
