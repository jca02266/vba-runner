/**
 * Dictionary Access 式（`!` 演算子）のテスト (§5.6.14)
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

console.log('--- Starting Dictionary Bang (!) Access Tests ---');

// --- 1. Dictionary から ! で値を読み取る ---
{
    const ev = evalVBA('');

    const mockDict = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
    };
    mockDict.__map__.set('Price', 100);
    mockDict.__map__.set('Name', 'Apple');

    ev.set('myDict', mockDict);

    assert.strictEqual(ev.evalExpression('myDict!Price'), 100, 'myDict!Price = 100');
    assert.strictEqual(ev.evalExpression('myDict!Name'), 'Apple', 'myDict!Name = "Apple"');
    console.log('[PASS] ! で Dictionary から値を読み取る');
}

// --- 2. ! と . のネストアクセス ---
{
    const ev = evalVBA('');

    const subObj = { value: 42 };
    const mockDict = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
    };
    mockDict.__map__.set('SubObj', subObj);

    ev.set('myDict', mockDict);

    assert.strictEqual(ev.evalExpression('myDict!SubObj.value'), 42, 'myDict!SubObj.value = 42');
    console.log('[PASS] ! と . のネストアクセス');
}

console.log('\n✅ Dictionary Access (!): 全テスト通過');
