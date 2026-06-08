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

async function testEvents() {
    console.log("--- Testing Events and WithEvents ---");
    
    const classCode = `
        Event MyEvent(val As Integer)
        Sub Trigger(v As Integer)
            RaiseEvent MyEvent(v)
        End Sub
    `;

    const moduleCode = `
        Dim WithEvents obj As Class1
        Sub TestEvents()
            Set obj = New Class1
            obj.Trigger 42
        End Sub
        Sub obj_MyEvent(val As Integer)
            Debug.Print "Event caught: " & val
        End Sub
    `;

    const logs: string[] = [];
    const evaluator = new Evaluator((msg) => logs.push(msg));
    
    // Parse Class
    const classLexer = new Lexer(classCode);
    const classParser = new Parser(classLexer.tokenize());
    const classProgram = classParser.parse();
    evaluator.registerClass('Class1', {
        type: 'ClassDeclaration',
        name: 'Class1',
        fields: classProgram.body.filter(s => s.type === 'VariableDeclaration') as any[],
        procedures: classProgram.body.filter(s => s.type === 'ProcedureDeclaration') as any[],
        body: classProgram.body
    } as any);

    // Parse Module
    const modLexer = new Lexer(moduleCode);
    const modParser = new Parser(modLexer.tokenize());
    const modProgram = modParser.parse();
    evaluator.evaluateModule(modProgram);
    evaluator.resolveIdentifiers([{ ast: modProgram, moduleName: '' }]);
    
    console.log("Running TestEvents...");
    evaluator.callProcedure('TestEvents', []);
    assert(logs.includes("Event caught: 42"), "Event should be raised and caught by WithEvents handler");

    console.log("--- All Event Tests Passed! ---");
}

testEvents().catch(err => {
    console.error(err);
    process.exit(1);
});
