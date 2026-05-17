import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runVba(source: string): Evaluator {
    const tokens = new Lexer(source).tokenize();
    const parser = new Parser(tokens);
    const ast = parser.parse();
    const ev = new Evaluator(() => {});
    ev.evaluate(ast);
    return ev;
}

// 1. Property Get/Let with Module Variable
{
    const source = `
        Dim internalValue
        Property Get MyValue()
            MyValue = internalValue
        End Property
        Property Let MyValue(v)
            internalValue = v
        End Property
    `;
    const ev = runVba(source);
    
    ev.callProcedure('MyValue', [42], 'let');
    assert.strictEqual(ev.callProcedure('MyValue', []), 42, 'Property Get after Let');
    
    // Assignment statement in eval
    ev.evalExpression('MyValue = 100'); // Note: currently evalExpression treats MyValue = 100 as comparison if possible
    // To ensure assignment, we can use a wrapper sub
    const source2 = `
        Sub TestAssignment()
            MyValue = 100
        End Sub
    `;
    const ev2 = runVba(source + source2);
    ev2.callProcedure('TestAssignment', []);
    assert.strictEqual(ev2.callProcedure('MyValue', []), 100, 'Property Let called via assignment statement');
}

// 2. Property Get/Set
{
    const source = `
        Dim internalObj
        Property Get MyObj()
            Set MyObj = internalObj
        End Property
        Property Set MyObj(o)
            Set internalObj = o
        End Property
    `;
    const ev = runVba(source);
    const mockObj = { name: 'test' };
    
    ev.callProcedure('MyObj', [mockObj], 'set');
    assert.strictEqual(ev.callProcedure('MyObj', []), mockObj, 'Property Get after Set');
}

// 3. Read-only and Write-only properties
{
    const source = `
        Property Get ReadOnly()
            ReadOnly = "Fixed"
        End Property
        Dim writeOnlyVal
        Property Let WriteOnly(v)
            writeOnlyVal = v
        End Property
        Function GetWriteOnly()
            GetWriteOnly = writeOnlyVal
        End Function
    `;
    const ev = runVba(source);
    
    assert.strictEqual(ev.callProcedure('ReadOnly', []), "Fixed", 'Read-only property');
    ev.callProcedure('WriteOnly', ["Changed"], 'let');
    assert.strictEqual(ev.callProcedure('GetWriteOnly', []), "Changed", 'Write-only property via function');
    
    try {
        ev.callProcedure('ReadOnly', ["Error"], 'let');
        assert.strictEqual(true, false, 'Should have thrown error for writing to read-only property');
    } catch (e: any) {
        assert.strictEqual(e.message.includes('ReadOnly'), true, 'Error message for missing Property Let');
    }
}

// 4. Set with non-object
{
    const source = `
        Sub TestSetError()
            Dim x
            Set x = 123
        End Sub
    `;
    const ev = runVba(source);
    try {
        ev.callProcedure('TestSetError', []);
        assert.strictEqual(true, false, 'Should have thrown error for Set x = 123');
    } catch (e: any) {
        assert.strictEqual(e.message.includes('Object required'), true, 'Error message: Object required');
    }
}

console.log('\n✅ Property Get/Let/Set: 全テスト通過');
