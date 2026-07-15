/**
 * TypeOf...Is 式のテスト (§5.6.7)
 *
 * 検証する挙動:
 *   - TypeOf <obj> Is Object: オブジェクト型か否か
 *   - TypeOf <obj> Is Dictionary / Collection: 組み込みクラスの判定
 *   - TypeOf <obj> Is <UserClass>: ユーザー定義型の判定
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

console.log('--- Starting TypeOf...Is Tests ---');

const ev = evalVBA('');

// モックオブジェクトを設定
const mockDict = { __isVbaDict__: true };
const mockColl = { __isVbaCollection__: true };
const mockUserObj = { __vbaTypeName__: 'MyClass' };

ev.set('myDict', mockDict);
ev.set('myColl', mockColl);
ev.set('myObj', mockUserObj);

// --- TypeOf ... Is Object ---
{
    assert.strictEqual(ev.evalExpression('TypeOf myDict Is Object'), -1, 'Dictionary は Object');
    assert.strictEqual(ev.evalExpression('TypeOf myObj Is Object'), -1, 'ユーザー定義型は Object');
    assert.strictEqual(ev.evalExpression('TypeOf 123 Is Object'), 0, '数値は Object ではない');
    console.log('[PASS] TypeOf ... Is Object');
}

// --- TypeOf ... Is Dictionary ---
{
    assert.strictEqual(ev.evalExpression('TypeOf myDict Is Dictionary'), -1, 'Dictionary 判定');
    assert.strictEqual(ev.evalExpression('TypeOf myColl Is Dictionary'), 0, 'Collection は Dictionary ではない');
    console.log('[PASS] TypeOf ... Is Dictionary');
}

// --- TypeOf ... Is Collection ---
{
    assert.strictEqual(ev.evalExpression('TypeOf myColl Is Collection'), -1, 'Collection 判定');
    console.log('[PASS] TypeOf ... Is Collection');
}

// --- TypeOf ... Is <UserClass> ---
{
    assert.strictEqual(ev.evalExpression('TypeOf myObj Is MyClass'), -1, '正しいユーザー定義型');
    assert.strictEqual(ev.evalExpression('TypeOf myObj Is OtherClass'), 0, '異なるユーザー定義型は False');
    console.log('[PASS] TypeOf ... Is <UserClass>');
}

// --- TypeOf ... Is <Library.ClassName> (ドット修飾型名) ---
{
    // パース自体が通ることを確認（以前はパースエラー）
    const mockAdoRs = { __vbaTypeName__: 'ADODB.Recordset' };
    ev.set('rs', mockAdoRs);
    assert.strictEqual(ev.evalExpression('TypeOf rs Is ADODB.Recordset'), -1, 'ドット修飾型名の判定');
    assert.strictEqual(ev.evalExpression('TypeOf rs Is ADODB.Connection'), 0, '異なるドット修飾型は False');
    assert.strictEqual(ev.evalExpression('TypeOf rs Is Object'), -1, 'ドット修飾型も Object として認識');

    // 3段階ドットもパースできること
    const ev2 = evalVBASingle('');
    const mockDeep = { __vbaTypeName__: 'A.B.C' };
    ev2.set('x', mockDeep);
    assert.strictEqual(ev2.evalExpression('TypeOf x Is A.B.C'), -1, '3段階ドット修飾型名');
    console.log('[PASS] TypeOf ... Is <Library.ClassName> (ドット修飾)');
}

console.log('\n✅ TypeOf...Is: 全テスト通過');
