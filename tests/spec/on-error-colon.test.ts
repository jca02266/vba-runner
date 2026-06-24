/**
 * `On Error Resume Next` / `On Error GoTo Label` の直後に `:` で文を連結した場合のテスト
 *
 * parser.ts の parseOnErrorStatement() は "Resume Next" 部分を読み取る際、
 * Newline/EOF までではなく `:` でも止まる必要があるが、その判定が抜けていたため
 * `On Error Resume Next: Debug.Print "x"` のように同一行に `:` で後続文を書くと、
 * 後続文が丸ごと On Error の label 文字列として読み込まれて消えてしまっていた
 * （`On Error GoTo Label: ...` も同様）。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

// --- 1. On Error Resume Next の直後に colon で後続文を書ける ---
{
    const lines: string[] = [];
    evalVBASingle(`
        Sub Test()
            On Error Resume Next: Debug.Print "a": Debug.Print "b"
        End Sub
        Test
    `, { onPrint: (s) => lines.push(s) });
    assert.deepStrictEqual(lines, ['a', 'b'], 'On Error Resume Next: の後続文が消えずに実行される');
}
console.log('[PASS] On Error Resume Next: 後続文が colon 連結でも実行される');

// --- 2. On Error GoTo Label の直後にも colon で後続文を書ける ---
{
    const lines: string[] = [];
    evalVBASingle(`
        Sub Test()
            On Error GoTo Handler: Debug.Print "a"
            Exit Sub
        Handler:
            Debug.Print "handled"
        End Sub
        Test
    `, { onPrint: (s) => lines.push(s) });
    assert.deepStrictEqual(lines, ['a'], 'On Error GoTo Label: の後続文が消えずに実行される');
}
console.log('[PASS] On Error GoTo Label: 後続文が colon 連結でも実行される');

// --- 3. On Error Resume Next: の後の文でエラーが起きても Resume Next が効く ---
{
    const lines: string[] = [];
    evalVBASingle(`
        Sub Test()
            On Error Resume Next: Err.Raise 1001, "S", "msg": Debug.Print "after"
        End Sub
        Test
    `, { onPrint: (s) => lines.push(s) });
    assert.deepStrictEqual(lines, ['after'], 'colon連結後のErr.RaiseもResume Nextで継続する');
}
console.log('[PASS] colon連結後の Err.Raise も Resume Next で継続する');

console.log('\n✅ On Error + colon 連結: 全テスト通過');
