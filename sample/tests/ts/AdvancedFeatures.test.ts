import { Evaluator, vbaTrue } from '../../../src/engine/evaluator';
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

async function testAdvanced() {
    console.log("--- Testing AddressOf and ReDim Preserve ---");
    
    const vbaCode = `
        Sub TestAdvanced()
            ' AddressOf
            Dim p As String
            p = AddressOf MySub
            Debug.Print "AddressOf: " & p
            
            ' ReDim Preserve
            Dim arr() As Integer
            ReDim arr(1)
            arr(0) = 10
            arr(1) = 20
            ReDim Preserve arr(2)
            arr(2) = 30
            Debug.Print "Arr(0): " & arr(0)
            Debug.Print "Arr(1): " & arr(1)
            Debug.Print "Arr(2): " & arr(2)
        End Sub

        Sub MySub()
        End Sub
    `;

    const logs: string[] = [];
    const evaluator = new Evaluator((msg) => logs.push(msg));
    
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    
    evaluator.evaluate(program);
    evaluator.callProcedure('TestAdvanced', []);
    
    assert(logs.includes("AddressOf: MySub"), "AddressOf should return procedure name");
    assert(logs.includes("Arr(0): 10"), "ReDim Preserve should keep existing data (0)");
    assert(logs.includes("Arr(1): 20"), "ReDim Preserve should keep existing data (1)");
    assert(logs.includes("Arr(2): 30"), "ReDim Preserve should add new data (2)");

    console.log("--- All Advanced Tests Passed! ---");
}

testAdvanced().catch(err => {
    console.error(err);
    process.exit(1);
});
