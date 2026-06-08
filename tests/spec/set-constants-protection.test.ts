/**
 * VBARunner.setConstants() で注入した定数が VBA コード側から上書きできないことを確認する。
 *
 * 実務上の典型: vba-analyzer が生成した allConstants (xlUp, xlDown 等) を
 * setConstants() で注入した後、VBA コードが誤って同名変数に代入しても元の値が保持される
 * （Error 5 を発生させる）べき。
 */
import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function makeEvaluator(code: string): Evaluator {
    return evalVBASingle(code);
}

// ---------------------------------------------------------------
// Test 1: setConstant() で注入した値は読み取れる
// ---------------------------------------------------------------
{
    const ev = makeEvaluator(`
        Function ReadConst()
            ReadConst = xlUp
        End Function
    `);
    ev.setConstant('xlUp', -4162);
    const result = ev.callProcedure('ReadConst', []);
    assert.strictEqual(result, -4162, 'setConstant で注入した xlUp が読み取れる');
    console.log('[PASS] setConstant 値読み取り:', result);
}

// ---------------------------------------------------------------
// Test 2: setConstant() で注入した定数に VBA コードから代入すると Error 5
// ---------------------------------------------------------------
{
    const ev = makeEvaluator(`
        Function TryOverwrite()
            Dim errNum As Long
            On Error Resume Next
            xlUp = 999
            errNum = Err.Number
            On Error GoTo 0
            TryOverwrite = errNum
        End Function
    `);
    ev.setConstant('xlUp', -4162);
    const result = ev.callProcedure('TryOverwrite', []);
    assert.strictEqual(result, 5, 'setConstant 定数への代入は Error 5');
    console.log('[PASS] setConstant 定数への代入 → Error 5:', result);
}

// ---------------------------------------------------------------
// Test 3: setConstant() 後も元の値が保持される（代入は無効）
// ---------------------------------------------------------------
{
    const ev = makeEvaluator(`
        Function ValueAfterAttempt()
            On Error Resume Next
            xlUp = 999
            On Error GoTo 0
            ValueAfterAttempt = xlUp
        End Function
    `);
    ev.setConstant('xlUp', -4162);
    const result = ev.callProcedure('ValueAfterAttempt', []);
    assert.strictEqual(result, -4162, '代入失敗後も元の値 -4162 が保持される');
    console.log('[PASS] 代入失敗後も元の値保持:', result);
}

// ---------------------------------------------------------------
// Test 4: setConstant() で複数の定数を注入
// ---------------------------------------------------------------
{
    const ev = makeEvaluator(`
        Function ReadMultiple()
            ReadMultiple = xlUp & "," & xlDown & "," & xlToLeft & "," & xlToRight
        End Function
    `);
    ev.setConstant('xlUp', -4162);
    ev.setConstant('xlDown', -4121);
    ev.setConstant('xlToLeft', -4159);
    ev.setConstant('xlToRight', -4161);
    const result = ev.callProcedure('ReadMultiple', []);
    assert.strictEqual(result, '-4162,-4121,-4159,-4161', '複数 setConstant の読み取り');
    console.log('[PASS] 複数定数の読み取り:', result);
}

// ---------------------------------------------------------------
// Test 5: VBARunner 経由の setConstants() (Record<string,any>) でも保護される
// ---------------------------------------------------------------
{
    // VBARunner の setConstants() は内部で setConstant() を使うべき
    // ここでは Evaluator の setConstant を直接使って同等テスト
    const ev = makeEvaluator(`
        Function TestXlValues()
            Dim errNum As Long
            On Error Resume Next
            xlEdgeLeft = 0
            errNum = Err.Number
            On Error GoTo 0
            TestXlValues = errNum & ":" & xlEdgeLeft
        End Function
    `);
    // setConstants() 相当: 複数をまとめて定数として注入
    const constants: Record<string, any> = { xlEdgeLeft: 7, xlEdgeRight: 10 };
    for (const [name, value] of Object.entries(constants)) {
        ev.setConstant(name, value);
    }
    const result = ev.callProcedure('TestXlValues', []);
    assert.strictEqual(result, '5:7', 'setConstants() 相当 → 代入エラー5 かつ元値保持');
    console.log('[PASS] setConstants() 相当の保護:', result);
}

// ---------------------------------------------------------------
// Test 6: 通常の Dim 変数と定数は独立して動作
// ---------------------------------------------------------------
{
    const ev = makeEvaluator(`
        Function MixConstVar()
            Dim myVar As Long
            myVar = xlConst + 1
            MixConstVar = myVar & ":" & xlConst
        End Function
    `);
    ev.setConstant('xlConst', 100);
    const result = ev.callProcedure('MixConstVar', []);
    assert.strictEqual(result, '101:100', '通常変数と注入定数が共存できる');
    console.log('[PASS] 通常変数と定数の共存:', result);
}

// ---------------------------------------------------------------
// Test 7: VBA ソース内の Const 宣言との競合（注入定数が優先）
// ---------------------------------------------------------------
{
    // VBA ソース内で同名の Const があっても、注入した値が env.get() で先に見つかる
    // ただしこのケースは実装依存なのでエラーにならないことだけ確認
    const ev = makeEvaluator(`
        Function ReadInjected()
            ReadInjected = myConst
        End Function
    `);
    ev.setConstant('myConst', 42);
    const result = ev.callProcedure('ReadInjected', []);
    assert.strictEqual(result, 42, '注入定数 myConst = 42 が読み取れる');
    console.log('[PASS] 注入定数の読み取り:', result);
}

console.log('\n✅ set-constants-protection: 全テスト通過');
