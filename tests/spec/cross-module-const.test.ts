/**
 * モジュールレベル定数のクロスモジュール参照と循環検出のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function makeModules(sources: Record<string, string>) {
    const ev = new Evaluator(console.log);
    const modules: Array<{ ast: any; moduleName: string }> = [];

    for (const [moduleName, src] of Object.entries(sources)) {
        ev.setSourceModule(moduleName);
        const ast = new Parser(new Lexer(src).tokenize()).parse();
        ev.evaluate(ast);
        modules.push({ ast, moduleName });
    }

    ev.reEvaluateModuleConstsAll(modules);
    return ev;
}

// --- 1. 非循環: ロード順に依存しない ---
{
    // Module2 を先にロードしても Module1 の定数を正しく参照できる
    const ev = makeModules({
        Module2: `Option Explicit\nPrivate Const myDir = xlUp`,
        Module1: `Public Const xlUp As Long = -4162`,
    });
    assert.strictEqual(ev.get('myDir'), -4162,
        '逆ロード順でもクロスモジュール定数が正しく解決される');
}

// --- 2. チェーン参照 ---
{
    const ev = makeModules({
        ModA: `Public Const BaseVal As Long = 10`,
        ModB: `Public Const Increment As Long = BaseVal + 5`,
        ModC: `Public Const Limit As Long = Increment * 2`,
    });
    assert.strictEqual(ev.get('limit'), 30,
        'チェーン参照: BaseVal=10, Increment=15, Limit=30');
}

// --- 3. 循環参照はエラー ---
{
    let threw = false;
    let msg = '';
    try {
        makeModules({
            M1: `Public Const ConstAlpha As Long = ConstBeta + 1`,
            M2: `Public Const ConstBeta As Long = ConstAlpha + 1`,
        });
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, '循環参照はエラーになる');
    assert.ok(msg.includes('Circular reference'), `エラーメッセージに "Circular reference" を含む: ${msg}`);
}

// --- 4. 自己参照も循環 ---
{
    let threw = false;
    try {
        makeModules({ M: `Public Const SelfRef As Long = SelfRef + 1` });
    } catch (e: any) {
        threw = true;
    }
    assert.ok(threw, '自己参照も循環として検出される');
}

// --- 5. Pass 1 後に未定義名が env に登録されていないこと ---
// このテストがなかったため Pass 1 の暗黙初期化バグを検出できなかった
{
    const ev = new Evaluator(console.log);
    // xlUp を定義しないまま xlUp を参照する const を含むモジュールをロード
    const ast = new Parser(new Lexer(`Private Const myDir = xlUp`).tokenize()).parse();
    ev.setSourceModule('Module1');
    ev.evaluate(ast);

    // Pass 1 終了時点では xlup は env に登録されていてはいけない
    assert.ok(!ev['env'].hasVariable('xlup'),
        'Pass 1 後: 未定義名 xlup が env に登録されていない');
    assert.ok(!ev['env'].hasVariable('mydir'),
        'Pass 1 後: 定数 mydir も env に登録されていない');

    // Pass 2 後も xlup はどこにも定義されていないので登録されない
    ev.reEvaluateModuleConstsAll([{ ast, moduleName: 'Module1' }]);
    assert.ok(!ev['env'].hasVariable('xlup'),
        'Pass 2 後: どこにも定義されていない xlup は env に存在しない');
}

// --- 6. 未定義名を手続き内で直接使うと Option Explicit が機能する ---
// const 右辺経由ではなく、手続き内で直接参照した場合のテスト
{
    const ev = new Evaluator(console.log);
    // xlUp を手続きの引数として渡す（引数位置の識別子はチェッカーが検査する）
    const src = [
        'Option Explicit',
        'Function Identity(x As Long) As Long',
        '  Identity = x',
        'End Function',
        'Function UseXlUp() As Long',
        '  UseXlUp = Identity(xlUp)',
        'End Function',
    ].join('\n');
    const ast = new Parser(new Lexer(src).tokenize()).parse();
    ev.setSourceModule('Module1');
    ev.evaluate(ast);
    ev.reEvaluateModuleConstsAll([{ ast, moduleName: 'Module1' }]);

    // xlUp はどこにも定義されていないので Option Explicit 違反
    let threw = false;
    try {
        ev.callProcedure('UseXlUp', []);
    } catch {
        threw = true;
    }
    assert.ok(threw, '手続き内で未定義の xlup を引数として使うと Option Explicit エラーになる');
}

console.log('cross-module-const: all tests passed');
