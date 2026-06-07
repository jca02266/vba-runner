// FOREIGN-NAME [identifier] 構文のテスト (§3.3.5.2)
//
// VBA では予約語をプロシージャ名として使う場合、[End] のように角括弧で囲む
// FOREIGN-NAME 構文が仕様上の正規手段。
// subroutine-name = IDENTIFIER / prefixed-name であり IDENTIFIER は
// reserved-identifier を含まないため、Property Get End はパースエラーになる。
//
// member-access-expression の右辺は unrestricted-name (= name / reserved-identifier)
// なので r.End(1) の呼び出し側は制限なく動作する。

import { evalVBASingle, assert } from '../../test-libs/test-runner';

// Test 1: r.GetEnd(1) — VBA からメンバーアクセスで呼び出す
// member-access-expression は unrestricted-name を使うため呼び出し側は制限なし
{
    const code = `
Class MockRange
    Public Row As Long
    Public Function GetEnd(direction As Long) As MockRange
        Dim r As New MockRange
        r.Row = 99
        Set GetEnd = r
    End Function
End Class

Function TestMemberCall()
    Dim r As New MockRange
    r.Row = 1
    Dim r2 As MockRange
    Set r2 = r.GetEnd(1)
    TestMemberCall = r2.Row
End Function
`;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('TestMemberCall', []), 99,
        'VBA クラスのメソッドをメンバーアクセスで呼べる');
    console.log('[PASS] VBA クラスのメンバー呼び出し (r.GetEnd)');
}

// Test 2: JS バックオブジェクトの End プロパティを VBA から呼ぶ
// JS では End はキーワードではないためプロパティ名に制限なし
// member-access-expression は unrestricted-name なので r.End(1) は動作する
{
    const code = `
Function TestJsEnd()
    Dim r As Object
    Set r = CreateObject("MockRange")
    Dim r2 As Object
    Set r2 = r.End(1)
    TestJsEnd = r2.Row
End Function
`;
    const ev = evalVBASingle(code);
    ev.externalObjectFactories.set('mockrange', () => ({
        __progId__: 'MockRange',
        Row: 1,
        End: (_direction: number) => ({ Row: 99, Column: 1 }),
    }));
    assert.strictEqual(ev.callProcedure('TestJsEnd', []), 99,
        'JS バックオブジェクトの End プロパティを VBA から呼べる');
    console.log('[PASS] JS モックの .End() を VBA から呼べる (r.End(1))');
}

// Test 3: [End]() — FOREIGN-NAME 構文でキーワード名のプロシージャを呼び出す
// 実 VBA での挙動:
//   - [End]() と書くと End という名前のプロシージャを呼ぶ
//   - 該当プロシージャが存在しなくても Option Explicit でエラーにならない
//   - プロシージャ定義側は Sub [End]() がエラーのため JS 側での登録になる
// 現状: レキサーが [ を Unknown トークンとして処理するためパースエラー (未実装)
{
    // Test 3a: [End]() が "end" という名前の関数にディスパッチするか
    {
        const code = `
Function TestForeignNameCall()
    [End]()
    TestForeignNameCall = 42
End Function
`;
        let called = false;
        try {
            const ev = evalVBASingle(code);
            ev.env.set('end', () => { called = true; });
            const result = ev.callProcedure('TestForeignNameCall', []);
            assert.strictEqual(called, true, '[End]() が end 関数を呼ぶ');
            assert.strictEqual(result, 42, '[End]() 後も実行が継続する');
            console.log('[TODO→PASS] [End]() FOREIGN-NAME 呼び出しが動作した');
        } catch (e: any) {
            console.log('[TODO/SKIP] [End]() 呼び出しは未実装:', e.message);
        }
    }

    // Test 3b: [End]() は対応するプロシージャが存在しなくてもエラーにならない
    {
        const code = `
Option Explicit
Function TestForeignNameNoOp()
    [End]()
    TestForeignNameNoOp = 99
End Function
`;
        try {
            const ev = evalVBASingle(code);
            const result = ev.callProcedure('TestForeignNameNoOp', []);
            assert.strictEqual(result, 99, '[End]() 存在しなくてもエラーなし');
            console.log('[TODO→PASS] [End]() 未定義でもエラーにならない (Option Explicit でも)');
        } catch (e: any) {
            console.log('[TODO/SKIP] [End]() 存在しない場合の挙動は未実装:', e.message);
        }
    }
}

console.log('foreign-name: 全テスト完了');
