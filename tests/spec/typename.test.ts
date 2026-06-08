import { vbaMissing } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// TypeName のテスト
{
    const code = `
        Function TestBasic()
            Debug.Print TypeName(10)
            Debug.Print TypeName(10.5)
            Debug.Print TypeName("hello")
            Debug.Print TypeName(True)
        End Function
        
        Function TestNull()
            TestNull = TypeName(Null)
        End Function
        
        Function TestNothing()
            TestNothing = TypeName(Nothing)
        End Function
        
        Function TestArray()
            Dim arr(5)
            TestArray = TypeName(arr)
        End Function

        Function TestMissing(Optional x)
            TestMissing = TypeName(x)
        End Function
    `;
    
    const ev = evalVBA(code);
    // Basic types
    assert.strictEqual(ev.evalExpression('TypeName(10)'), 'Integer', 'TypeName(10) -> Integer');
    assert.strictEqual(ev.evalExpression('TypeName("abc")'), 'String', 'TypeName("abc") -> String');
    assert.strictEqual(ev.evalExpression('TypeName(True)'), 'Boolean', 'TypeName(True) -> Boolean');

    // Null/Nothing/Empty
    assert.strictEqual(runFunc(code, 'TestNull'), 'Null', 'TypeName(Null) -> Null');
    assert.strictEqual(runFunc(code, 'TestNothing'), 'Nothing', 'TypeName(Nothing) -> Nothing');
    
    // Array
    assert.strictEqual(runFunc(code, 'TestArray'), 'Variant()', 'TypeName(Array) -> Variant()');

    // Missing
    assert.strictEqual(runFunc(code, 'TestMissing', []), 'Error', 'TypeName(Missing) -> Error');

    console.log('[PASS] TypeName の基本動作');
}

console.log('\n✅ TypeName: 全テスト通過');
