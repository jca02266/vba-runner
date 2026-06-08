/**
 * モジュールレベル定数のクロスモジュール参照と循環検出のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

// --- 1. 非循環: ロード順に依存しない（Public Const はクロスモジュール参照可）---
{
    const ev = evalVBAModules([
        { name: 'Module2', code: 'Option Explicit\nPublic Const myDir = xlUp' },
        { name: 'Module1', code: 'Public Const xlUp As Long = -4162' },
    ]);
    assert.strictEqual(ev.get('myDir'), -4162,
        '逆ロード順でもクロスモジュール Public Const が正しく解決される');
}

// --- 2. チェーン参照 ---
{
    const ev = evalVBAModules([
        { name: 'ModA', code: 'Public Const BaseVal As Long = 10' },
        { name: 'ModB', code: 'Public Const Increment As Long = BaseVal + 5' },
        { name: 'ModC', code: 'Public Const Limit As Long = Increment * 2' },
    ]);
    assert.strictEqual(ev.get('limit'), 30,
        'チェーン参照: BaseVal=10, Increment=15, Limit=30');
}

// --- 3. クロスモジュール循環参照はエラー（クロスモジュールとして先に検出される）---
{
    let threw = false;
    try {
        evalVBAModules([
            { name: 'M1', code: 'Public Const ConstAlpha As Long = ConstBeta + 1' },
            { name: 'M2', code: 'Public Const ConstBeta As Long = ConstAlpha + 1' },
        ]);
    } catch {
        threw = true;
    }
    assert.ok(threw, 'クロスモジュール循環参照はエラーになる');
}

// --- 3b. 同一モジュール内の循環参照は "Circular reference" エラー ---
{
    let threw = false;
    let msg = '';
    try {
        evalVBAModules([
            { name: 'M', code: 'Public Const A As Long = B + 1\nPublic Const B As Long = A + 1' },
        ]);
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, '同一モジュール内循環参照はエラーになる');
    assert.ok(msg.includes('Circular reference'), `エラーメッセージに "Circular reference" を含む: ${msg}`);
}

// --- 4. 自己参照も循環 ---
{
    let threw = false;
    try {
        evalVBAModules([{ name: 'M', code: 'Public Const SelfRef As Long = SelfRef + 1' }]);
    } catch (e: any) {
        threw = true;
    }
    assert.ok(threw, '自己参照も循環として検出される');
}

// --- 5. 未定義名を参照する Const は "Constant expression required" エラーになる ---
{
    const ev = new Evaluator(console.log);
    const ast = new Parser(new Lexer(`Private Const myDir = xlUp`).tokenize()).parse();
    ev.setSourceModule('Module1');
    ev.evaluateModule(ast);

    // Pass 1 終了時点では xlup も mydir も env に登録されていない
    assert.ok(!ev['env'].hasVariable('xlup'),
        'Pass 1 後: 未定義名 xlup が env に登録されていない');
    assert.ok(!ev['env'].hasVariable('mydir'),
        'Pass 1 後: 定数 mydir も env に登録されていない');

    // Pass 2 で xlup が未定義のため "Constant expression required" エラーになる
    let threw = false;
    let msg = '';
    try {
        ev.resolveIdentifiers([{ ast, moduleName: 'Module1' }]);
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, 'Pass 2: 未定義名を参照する Const はエラーになる');
    assert.ok(msg.includes('Constant expression required'),
        `エラーメッセージに "Constant expression required" を含む: ${msg}`);
}

// --- 6. 未定義名を手続き内で直接使うと Option Explicit が機能する ---
// const 右辺経由ではなく、手続き内で直接参照した場合のテスト
{
    const src = [
        'Option Explicit',
        'Function Identity(x As Long) As Long',
        '  Identity = x',
        'End Function',
        'Function UseXlUp() As Long',
        '  UseXlUp = Identity(xlUp)',
        'End Function',
    ].join('\n');
    const ev = evalVBASingle(src);

    // xlUp はどこにも定義されていないので Option Explicit 違反
    let threw = false;
    try {
        ev.callProcedure('UseXlUp', []);
    } catch {
        threw = true;
    }
    assert.ok(threw, '手続き内で未定義の xlup を引数として使うと Option Explicit エラーになる');
}

// --- 7. Private Const は他モジュールから参照不可 ---
{
    let threw = false;
    let msg = '';
    try {
        evalVBAModules([
            { name: 'Module1', code: 'Public Const A = B' },
            { name: 'Module2', code: 'Private Const B As Long = 42' },
        ]);
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, '他モジュールの Private Const 参照はエラーになる');
    assert.ok(msg.includes('Constant expression required'),
        `エラーメッセージに "Constant expression required" を含む: ${msg}`);
}

// --- 8. Public Const は他モジュールから参照可 ---
{
    const ev = evalVBAModules([
        { name: 'Module1', code: 'Public Const A = B' },
        { name: 'Module2', code: 'Public Const B As Long = 42' },
    ]);
    assert.strictEqual(ev.get('a'), 42, '他モジュールの Public Const は参照可');
}

console.log('cross-module-const: all tests passed');
