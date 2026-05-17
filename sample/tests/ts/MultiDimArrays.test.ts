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

async function testMultiDimArrays() {
    console.log("--- Testing Multi-Dimensional Arrays ---");
    
    const vbaCode = `
        Sub TestMultiDim()
            Dim arr(1 To 3, 2 To 4) As Integer
            
            arr(1, 2) = 10
            arr(3, 4) = 20
            
            Debug.Print "arr(1, 2): " & arr(1, 2)
            Debug.Print "arr(3, 4): " & arr(3, 4)
            Debug.Print "LBound1: " & LBound(arr, 1)
            Debug.Print "UBound1: " & UBound(arr, 1)
            Debug.Print "LBound2: " & LBound(arr, 2)
            Debug.Print "UBound2: " & UBound(arr, 2)
            
            ReDim Preserve arr(1 To 3, 2 To 5)
            arr(3, 5) = 30
            
            Debug.Print "Preserve arr(1, 2): " & arr(1, 2)
            Debug.Print "Preserve arr(3, 5): " & arr(3, 5)
            Debug.Print "Preserve UBound2: " & UBound(arr, 2)
        End Sub
    `;

    const logs: string[] = [];
    const evaluator = new Evaluator((msg) => logs.push(msg));
    
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    
    evaluator.evaluate(program);
    evaluator.callProcedure('TestMultiDim', []);
    
    assert(logs.includes("arr(1, 2): 10"), "Basic 2D assignment (1,2)");
    assert(logs.includes("arr(3, 4): 20"), "Basic 2D assignment (3,4)");
    assert(logs.includes("LBound1: 1"), "LBound dimension 1 is 1");
    assert(logs.includes("UBound1: 3"), "UBound dimension 1 is 3");
    assert(logs.includes("LBound2: 2"), "LBound dimension 2 is 2");
    assert(logs.includes("UBound2: 4"), "UBound dimension 2 is 4");
    
    assert(logs.includes("Preserve arr(1, 2): 10"), "ReDim Preserve keeps data");
    assert(logs.includes("Preserve arr(3, 5): 30"), "ReDim Preserve allows new data");
    assert(logs.includes("Preserve UBound2: 5"), "ReDim Preserve updates UBound of last dimension");

    console.log("--- All Multi-Dimensional Array Tests Passed! ---");
}

testMultiDimArrays().catch(err => {
    console.error(err);
    process.exit(1);
});
