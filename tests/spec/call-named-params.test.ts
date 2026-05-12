/**
 * Call Statement (§5.4.2.1) & Named Parameters (§5.6.3.1.2) のテスト
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

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

// --- 1. Call Statement の基本動作 ---
const callCode = `
    Dim globalVal
    Sub SetVal(v)
        globalVal = v
    End Sub

    Sub Main()
        Call SetVal(100)
    End Sub
`;
const ev1 = evalVBA(callCode);
ev1.callProcedure('Main', []);
assert.strictEqual(ev1.env.get('globalval'), 100, 'Call SetVal(100) で変数が更新される');
console.log('[PASS] Call Statement 基本動作');

// --- 2. Named Parameters (名前付き引数) ---
const namedParamCode = `
    Function Subtract(a, b)
        Subtract = a - b
    End Function

    Function TestNamed()
        TestNamed = Subtract(b:=10, a:=50)
    End Function
`;
assert.strictEqual(runFunc(namedParamCode, 'TestNamed'), 40, 'b:=10, a:=50 で 50-10=40');
console.log('[PASS] Named Parameters 基本動作');

// --- 3. 名前付き引数と通常引数の混在 ---
const mixedCode = `
    Function Mixed(a, b, Optional c = 0)
        Mixed = a + b + c
    End Function

    Function TestMixed()
        TestMixed = Mixed(10, c:=30, b:=20)
    End Function
`;
assert.strictEqual(runFunc(mixedCode, 'TestMixed'), 60, '10 + 20 + 30 = 60');
console.log('[PASS] 通常引数と名前付き引数の混在');

// --- 4. Call と Named Params の組み合わせ ---
const callNamedCode = `
    Dim result
    Sub Store(val)
        result = val
    End Sub

    Sub Main()
        Call Store(val:=999)
    End Sub
`;
const ev4 = evalVBA(callNamedCode);
ev4.callProcedure('Main', []);
assert.strictEqual(ev4.env.get('result'), 999, 'Call Store(val:=999) が動作する');
console.log('[PASS] Call + Named Parameters');

console.log('\n✅ Call & Named Parameters: 全テスト通過');
