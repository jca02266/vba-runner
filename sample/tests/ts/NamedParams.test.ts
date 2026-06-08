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

async function testNamedParams() {
    console.log("--- Testing Named Parameters and Attributes ---");
    
    const vbaCode = `
        Attribute VB_Name = "MyModule"
        
        Sub TestCall()
            Greet Name:="Koji", Greeting:="Hello"
        End Sub
        
        Sub Greet(Greeting As String, Name As String)
            Debug.Print Greeting & ", " & Name
        End Sub
    `;

    const logs: string[] = [];
    const evaluator = new Evaluator((msg) => logs.push(msg));
    
    const lexer = new Lexer(vbaCode);
    const parser = new Parser(lexer.tokenize());
    const program = parser.parse();
    
    evaluator.evaluateModule(program);
    evaluator.resolveIdentifiers([{ ast: program, moduleName: '' }]);
    
    console.log("Running TestCall...");
    evaluator.callProcedure('TestCall', []);
    assert(logs.includes("Hello, Koji"), "Named parameters should be mapped correctly regardless of order");

    console.log("--- All Tests Passed! ---");
}

testNamedParams().catch(err => {
    console.error(err);
    process.exit(1);
});
