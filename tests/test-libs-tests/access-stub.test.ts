/**
 * Access API スタブモード（injectAccessStub）のテスト
 */
import { VBARunner, assert } from '../../test-libs/test-runner';
import { injectAccessStub } from '../../test-libs/access-stub';

const runner = new VBARunner('tests/fixtures/access-stub');
const app = injectAccessStub(runner.evaluator);

// --- DoCmd ノーオプが例外なし ---
assert.strictEqual(runner.run('TestDoCmd', []), 'ok', 'DoCmd ノーオプ');
console.log('[PASS] DoCmd.SetWarnings / RunSQL / OpenForm / Close');

// --- 空 Recordset の EOF = True ---
assert.isTrue(runner.run('TestEmptyRecordset', []), '空 Recordset.EOF = True');
console.log('[PASS] 空 Recordset EOF');

// --- Recordset にデータを注入して読む ---
app.db.setTableData('Customers', [{ Name: 'Alice', Age: 30 }]);
assert.strictEqual(runner.run('TestRecordsetRead', []), 'Alice', 'Recordset Fields("Name").Value');
console.log('[PASS] Recordset Fields 読み取り');

// --- ループで件数をカウント ---
app.db.setTableData('Orders', [
    { OrderId: 1, Amount: 100 },
    { OrderId: 2, Amount: 200 },
    { OrderId: 3, Amount: 300 },
]);
assert.strictEqual(runner.run('TestRecordsetLoop', []), 3, 'Recordset ループ 3件');
console.log('[PASS] Recordset MoveNext ループ');

// --- Forms("name").setControl でコントロール値を事前登録 ---
app.Forms('frmCustomer').setControl('txtName', 'Bob');
assert.strictEqual(runner.run('TestFormControl', []), 'Bob', 'Forms フォームコントロール');
console.log('[PASS] Forms("frmCustomer").txtName.Value');

// --- Application プロパティ ---
assert.strictEqual(runner.run('TestAppProps', []), 'Microsoft Access', 'Application.Name');
console.log('[PASS] Application.Echo / ScreenUpdating / Name');

console.log('\n✅ access-stub: 全テスト通過');
