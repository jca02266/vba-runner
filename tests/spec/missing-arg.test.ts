import { Evaluator, VbaErrorCode } from '../../src/engine/evaluator';
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

function expectVbaError(code: string, expectedNumber: number, label: string): void {
    let caught: any = null;
    try {
        evalVBA(code);
    } catch (e: any) {
        caught = e;
    }
    assert.strictEqual(caught !== null, true, `${label}: エラーが発生する`);
    assert.strictEqual(caught?.number, expectedNumber, `${label}: エラー番号 ${expectedNumber}（got: ${caught?.number}）`);
}

// 1. 中間引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Add3(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Add3 = a + b + c
        End Function
        Function Test()
            Test = Add3(10,,30)
        End Function
    `, 'Test');
    assert.strictEqual(r, 42, 'Add3(10,,30) = 10+2+30 = 42');
    console.log('[PASS] 中間省略: デフォルト値適用:', r);
}

// 2. 末尾引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Greet(name As String, Optional greeting As String = "Hello")
            Greet = greeting & " " & name
        End Function
        Function Test()
            Test = Greet("World",)
        End Function
    `, 'Test');
    assert.strictEqual(r, 'Hello World', 'Greet("World",) = "Hello World"');
    console.log('[PASS] 末尾省略: デフォルト値適用:', r);
}

// 3. 先頭引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Test()
            Test = Func(,20)
        End Function
        Function Func(Optional a As Long = 10, Optional b As Long = 0)
            Func = a + b
        End Function
    `, 'Test');
    assert.strictEqual(r, 30, 'Func(,20) = 10+20 = 30');
    console.log('[PASS] 先頭省略: デフォルト値適用:', r);
}

// 4. IsMissing は省略スロットで True になる（デフォルト値なし Optional）
{
    const r = runFunc(`
        Function Test()
            Test = Check(1,,3)
        End Function
        Function Check(Optional a, Optional b, Optional c)
            If IsMissing(b) Then
                Check = "missing"
            Else
                Check = "present"
            End If
        End Function
    `, 'Test');
    assert.strictEqual(r, 'missing', 'IsMissing(b) = True when ,, used');
    console.log('[PASS] 省略スロットで IsMissing = True:', r);
}

// 5. 省略なしの通常呼び出しは従来通り
{
    const r = runFunc(`
        Function Add3(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Add3 = a + b + c
        End Function
        Function Test()
            Test = Add3(10, 20, 30)
        End Function
    `, 'Test');
    assert.strictEqual(r, 60, 'Add3(10,20,30) = 60');
    console.log('[PASS] 省略なし通常呼び出し:', r);
}

// 6. キーワード引数と省略スロットの混在
{
    const r = runFunc(`
        Function Func(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Func = a * 100 + b * 10 + c
        End Function
        Function Test()
            Test = Func(,, c:=9)
        End Function
    `, 'Test');
    assert.strictEqual(r, 129, 'Func(,,c:=9) = 1*100+2*10+9 = 129');
    console.log('[PASS] 省略スロット + キーワード引数:', r);
}

// 7. 省略スロットが必須パラメーター（Optional でなく defaultValue もない）の位置に来ると
//    449 (Argument not optional) になる（標準モジュールの Sub/Function）
{
    expectVbaError(`
        Sub Foo(a As Long, b As Long, c As Long)
        End Sub
        Call Foo(1, , 3)
    `, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Foo(1, , 3) の b は必須');
    console.log('[PASS] 必須パラメーターの省略スロットは 449（標準プロシージャ）');
}

// 8. 同じ状況でクラスメソッド呼び出しでも 449 になる
{
    expectVbaError(`
        Class Foo
            Public Sub Bar(a As Long, b As Long, c As Long)
            End Sub
        End Class
        Dim f As New Foo
        Call f.Bar(1, , 3)
    `, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'f.Bar(1, , 3) の b は必須');
    console.log('[PASS] 必須パラメーターの省略スロットは 449（クラスメソッド）');
}

// 9. Optional パラメーターの省略スロットは引き続きエラーにならない（既存挙動の確認）
{
    const r = runFunc(`
        Function Func(a As Long, Optional b As Long = 2, Optional c As Long = 3)
            Func = a * 100 + b * 10 + c
        End Function
        Function Test()
            Test = Func(1, , 9)
        End Function
    `, 'Test');
    assert.strictEqual(r, 129, 'Func(1, , 9) = 1*100+2*10+9 = 129（Optional の省略は従来通り）');
    console.log('[PASS] Optional パラメーターの省略スロットはエラーにならない:', r);
}

// 10. 同じ状況でモジュール修飾呼び出し（Module1.Proc(...)）でも 449 になる
{
    let caught: any = null;
    try {
        evalVBAModules([
            { name: 'Module1', code: 'Sub Foo(a As Long, b As Long, c As Long)\nEnd Sub' },
            { name: 'Main', code: 'Sub Test()\nCall Module1.Foo(1, , 3)\nEnd Sub' },
        ]).callProcedure('Test', []);
    } catch (e: any) {
        caught = e;
    }
    assert.strictEqual(caught !== null, true, 'Module1.Foo(1, , 3) の b は必須: エラーが発生する');
    assert.strictEqual(caught?.number, VbaErrorCode.ARGUMENT_NOT_OPTIONAL,
        `Module1.Foo(1, , 3) の b は必須: エラー番号 449（got: ${caught?.number}）`);
    console.log('[PASS] 必須パラメーターの省略スロットは 449（モジュール修飾呼び出し）');
}

// 11. late-bound Dictionary/Collection の必須メンバーは、裸参照・空括弧で
//     JavaScript関数値や既定値を漏出させず、引数不足契約を適用する
{
    const code = String.raw`
Function ProbeDynamicMembers() As String
    Dim d As Object, c As Object, fso As Object, stream As Object, value As Variant, out As String
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "key", True
    Set c = New Collection
    c.Add "entry"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile("dynamic-read.txt", True, False)
    stream.Write "abc": stream.Close
    Set stream = fso.OpenTextFile("dynamic-read.txt", 1, False, 0)
    On Error Resume Next
    value = Empty: Err.Clear: value = d.Exists: out = "DICT-BARE=" & CStr(Err.Number) & ":" & TypeName(value)
    value = Empty: Err.Clear: value = d.Exists(): out = out & "|DICT-PARENS=" & CStr(Err.Number) & ":" & TypeName(value)
    Err.Clear: value = d.Exists("key"): out = out & "|DICT-KEY=" & CStr(Err.Number) & ":" & TypeName(value) & ":" & CStr(value)
    value = Empty: Err.Clear: value = c.Item: out = out & "|COLL-BARE=" & CStr(Err.Number) & ":" & TypeName(value)
    value = Empty: Err.Clear: value = c.Item(): out = out & "|COLL-PARENS=" & CStr(Err.Number) & ":" & TypeName(value)
    Err.Clear: value = c.Item(1): out = out & "|COLL-INDEX=" & CStr(Err.Number) & ":" & TypeName(value) & ":" & CStr(value)
    With d
        value = Empty: Err.Clear: value = .Exists: out = out & "|WITH-DICT=" & CStr(Err.Number) & ":" & TypeName(value)
    End With
    With c
        value = Empty: Err.Clear: value = .Item(): out = out & "|WITH-COLL=" & CStr(Err.Number) & ":" & TypeName(value)
    End With
    value = Empty: Err.Clear: value = stream.Read: out = out & "|READ-BARE=" & CStr(Err.Number) & ":" & TypeName(value)
    value = Empty: Err.Clear: value = stream.Read(): out = out & "|READ-PARENS=" & CStr(Err.Number) & ":" & TypeName(value)
    stream.Close
    ProbeDynamicMembers = out
End Function
`;
    const ev = evalVBA(code);
    const result = ev.callProcedure('ProbeDynamicMembers', []);
    assert.strictEqual(result,
        'DICT-BARE=450:Empty|DICT-PARENS=450:Empty|DICT-KEY=0:Boolean:True|COLL-BARE=450:Empty|COLL-PARENS=450:Empty|COLL-INDEX=0:String:entry|WITH-DICT=450:Empty|WITH-COLL=450:Empty|READ-BARE=450:Empty|READ-PARENS=450:Empty',
        'late-bound Dictionary/Collection の必須メンバー引数契約');
    console.log('[PASS] late-bound Dictionary/Collection required-member matrix');
}

// 12. FSO.GetFile も戻り値の型（Set/Variant）にかかわらず必須引数を検査する
{
    const code = String.raw`
Function ProbeFsoGetFile() As String
    Dim fso As Object, file As Object, value As Variant, out As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = Nothing: Err.Clear: Set file = fso.GetFile: out = "SET-BARE=" & CStr(Err.Number) & ":" & TypeName(file)
    Set file = Nothing: Err.Clear: Set file = fso.GetFile(): out = out & "|SET-PARENS=" & CStr(Err.Number) & ":" & TypeName(file)
    value = Empty: Err.Clear: value = fso.GetFile: out = out & "|VALUE-BARE=" & CStr(Err.Number) & ":" & TypeName(value)
    value = Empty: Err.Clear: value = fso.GetFile(): out = out & "|VALUE-PARENS=" & CStr(Err.Number) & ":" & TypeName(value)
    ProbeFsoGetFile = out
End Function
`;
    const ev = evalVBA(code);
    assert.strictEqual(ev.callProcedure('ProbeFsoGetFile', []),
        'SET-BARE=450:Nothing|SET-PARENS=450:Nothing|VALUE-BARE=450:Empty|VALUE-PARENS=450:Empty',
        'FSO.GetFile required pathname is validated in Set and Variant paths');
    console.log('[PASS] FSO.GetFile required-member matrix');
}

console.log('\n✅ missing-arg: 全テスト通過');
