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

// 数値リテラルのサフィックス型情報保持
{
    const ev = evalVBASingle('Function Dummy(): End Function');
    assert.strictEqual(ev.evalExpression('TypeName(100%)'), 'Integer', 'TypeName(100%) -> Integer');
    assert.strictEqual(ev.evalExpression('TypeName(100&)'), 'Long', 'TypeName(100&) -> Long');
    assert.strictEqual(ev.evalExpression('TypeName(1.5!)'), 'Single', 'TypeName(1.5!) -> Single');
    assert.strictEqual(ev.evalExpression('TypeName(1.5#)'), 'Double', 'TypeName(1.5#) -> Double');
    assert.strictEqual(ev.evalExpression('TypeName(1.5@)'), 'Currency', 'TypeName(1.5@) -> Currency');
    assert.strictEqual(ev.evalExpression('TypeName(123^)'), 'LongLong', 'TypeName(123^) -> LongLong');

    // VarType も確認
    assert.strictEqual(ev.evalExpression('VarType(100%)'), 2, 'VarType(100%) -> 2 (vbInteger)');
    assert.strictEqual(ev.evalExpression('VarType(100&)'), 3, 'VarType(100&) -> 3 (vbLong)');
    assert.strictEqual(ev.evalExpression('VarType(1.5!)'), 4, 'VarType(1.5!) -> 4 (vbSingle)');
    assert.strictEqual(ev.evalExpression('VarType(123^)'), 20, 'VarType(123^) -> 20 (vbLongLong)');

    // 小数点付きリテラルは isFloat により Double
    assert.strictEqual(ev.evalExpression('TypeName(1.0)'), 'Double', 'TypeName(1.0) -> Double');
    assert.strictEqual(ev.evalExpression('TypeName(1.5)'), 'Double', 'TypeName(1.5) -> Double');

    console.log('[PASS] 数値リテラルのサフィックス型情報保持');
}

// サフィックス付きリテラルのオーバーフロー検出
{
    function overflows(code: string): boolean {
        try {
            evalVBASingle(code).callProcedure('F', []);
            return false;
        } catch (e: any) {
            return /Overflow/.test(e.message);
        }
    }
    assert.strictEqual(overflows('Function F(): F = 100000%: End Function'), true, '100000% -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 40000%: End Function'), true, '40000% (>32767) -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 3000000000&: End Function'), true, '3000000000& -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 100&: End Function'), false, '100& -> no overflow');
    assert.strictEqual(overflows('Function F(): F = 32767%: End Function'), false, '32767% -> no overflow');

    console.log('[PASS] サフィックス付きリテラルのオーバーフロー検出');
}

console.log('\n✅ TypeName: 全テスト通過');
