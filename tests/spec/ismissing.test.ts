import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// IsMissing のテスト
{
    const code = `
        Function TestMissing(Optional x As Variant)
            TestMissing = IsMissing(x)
        End Function
        
        Function TestNotMissing(Optional x As Variant)
            TestNotMissing = IsMissing(x)
        End Function
    `;
    
    // 引数なしで呼び出し
    assert.strictEqual(runFunc(code, 'TestMissing', []), vbaTrue, 'IsMissing: should return True when argument is omitted');
    
    // 引数ありで呼び出し
    assert.strictEqual(runFunc(code, 'TestNotMissing', [10]), vbaFalse, 'IsMissing: should return False when argument is provided');
    assert.strictEqual(runFunc(code, 'TestNotMissing', [null]), vbaFalse, 'IsMissing: should return False when Null is provided');
    assert.strictEqual(runFunc(code, 'TestNotMissing', [""]), vbaFalse, 'IsMissing: should return False when empty string is provided');

    console.log('[PASS] IsMissing の基本動作');
}

console.log('\n✅ IsMissing: 全テスト通過');
