/**
 * VBScript.RegExp モック実装のテスト
 *
 * test-libs/regexp-mock.ts の `createRegExpMock` を登録し、
 * VBA から `CreateObject("VBScript.RegExp")` を呼んだ際に意図通り動作することを検証する。
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';
import { createRegExpMock } from '../../test-libs/regexp-mock';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.registerComObject(createRegExpMock);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

const helpers = `
    Function NewRe(pattern As String, Optional isGlobal As Boolean = True, Optional ignoreCase As Boolean = False) As Object
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = pattern
        re.Global = isGlobal
        re.IgnoreCase = ignoreCase
        Set NewRe = re
    End Function
`;

console.log('--- Starting VBScript.RegExp Mock Tests ---');

// --- 1. Test メソッド ---
{
    const code = helpers + `
        Function TestMatch(p As String, s As String) As Boolean
            TestMatch = NewRe(p).Test(s)
        End Function
    `;
    assert.isTrue(runFunc(code, 'TestMatch', ['\\d+', 'abc 123']), 'Test: 数字を含む → True');
    assert.isFalse(runFunc(code, 'TestMatch', ['\\d+', 'abc']), 'Test: 数字なし → False');
    console.log('[PASS] Test メソッド');
}

// --- 2. Execute: Count と Item ---
{
    const code = helpers + `
        Function CountMatches(s As String) As Long
            CountMatches = NewRe("\\d+").Execute(s).Count
        End Function

        Function FirstMatch(s As String) As String
            FirstMatch = NewRe("\\d+").Execute(s).Item(0).Value
        End Function
    `;
    assert.strictEqual(runFunc(code, 'CountMatches', ['1 22 333']), 3, 'Count = 3');
    assert.strictEqual(runFunc(code, 'CountMatches', ['abc']), 0, 'Count = 0');
    assert.strictEqual(runFunc(code, 'FirstMatch', ['abc 123 def']), '123', 'Item(0).Value');
    console.log('[PASS] Execute / Count / Item');
}

// --- 3. For Each で MatchCollection を反復 ---
{
    const code = helpers + `
        Function ConcatMatches(s As String) As String
            Dim re As Object
            Set re = NewRe("\\d+")
            Dim m As Object
            Dim result As String
            For Each m In re.Execute(s)
                result = result & m.Value & ","
            Next
            ConcatMatches = result
        End Function
    `;
    assert.strictEqual(runFunc(code, 'ConcatMatches', ['1 22 333']), '1,22,333,', 'For Each で反復');
    console.log('[PASS] For Each 反復');
}

// --- 4. Match の FirstIndex / Length ---
{
    const code = helpers + `
        Function FirstIdx(s As String) As Long
            FirstIdx = NewRe("\\d+").Execute(s).Item(0).FirstIndex
        End Function

        Function FirstLen(s As String) As Long
            FirstLen = NewRe("\\d+").Execute(s).Item(0).Length
        End Function
    `;
    assert.strictEqual(runFunc(code, 'FirstIdx', ['abc 123 def']), 4, 'FirstIndex (0-based)');
    assert.strictEqual(runFunc(code, 'FirstLen', ['abc 123 def']), 3, 'Length');
    console.log('[PASS] FirstIndex / Length');
}

// --- 5. SubMatches（キャプチャグループ）---
{
    const code = helpers + `
        Function GetGroup(s As String, idx As Long) As String
            Dim re As Object
            Set re = NewRe("(\\d+)-(\\d+)")
            GetGroup = re.Execute(s).Item(0).SubMatches(idx)
        End Function
    `;
    assert.strictEqual(runFunc(code, 'GetGroup', ['code: 12-345', 0]), '12', 'SubMatches(0)');
    assert.strictEqual(runFunc(code, 'GetGroup', ['code: 12-345', 1]), '345', 'SubMatches(1)');
    console.log('[PASS] SubMatches');
}

// --- 6. IgnoreCase ---
{
    const code = helpers + `
        Function CountCI(p As String, s As String) As Long
            CountCI = NewRe(p, True, True).Execute(s).Count
        End Function
    `;
    assert.strictEqual(runFunc(code, 'CountCI', ['hello', 'Hello HELLO hello']), 3, 'IgnoreCase=True で 3 マッチ');
    console.log('[PASS] IgnoreCase');
}

// --- 7. Replace ---
{
    const code = helpers + `
        Function MaskNumbers(s As String) As String
            MaskNumbers = NewRe("\\d+").Replace(s, "#")
        End Function
    `;
    assert.strictEqual(runFunc(code, 'MaskNumbers', ['abc 123 def 456']), 'abc # def #', 'Replace で全数字を #');
    console.log('[PASS] Replace');
}

// --- 8. Global = False（Execute は先頭 1 件のみ）---
{
    const code = helpers + `
        Function CountWithGlobalFalse(s As String) As Long
            Dim re As Object
            Set re = NewRe("\\d+", False)
            CountWithGlobalFalse = re.Execute(s).Count
        End Function
    `;
    assert.strictEqual(runFunc(code, 'CountWithGlobalFalse', ['1 22 333']), 1, 'Global=False は 1 件のみ');
    console.log('[PASS] Global=False');
}

// --- 9. 参照設定相当: `New RegExp` / `Dim re As RegExp` ---
// registerComObject(createRegExpMock) 1 回の登録で、
// `CreateObject("VBScript.RegExp")` だけでなく、
// `Set re = New RegExp` / `Dim re As RegExp` も同じ factory で生成される。
{
    const code = `
        Function CountViaNew(s As String) As Long
            Dim re As RegExp
            Set re = New RegExp
            re.Pattern = "\\d+"
            re.Global = True
            CountViaNew = re.Execute(s).Count
        End Function
    `;
    assert.strictEqual(runFunc(code, 'CountViaNew', ['1 22 333']), 3, 'New RegExp で 3 マッチ');
    console.log('[PASS] 参照設定相当 (New RegExp)');
}

// --- 10. CreateObject と New が同じ動作 ---
{
    const code = `
        Function MatchByCreateObject(s As String) As Boolean
            Dim re As Object
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "[a-z]+"
            MatchByCreateObject = re.Test(s)
        End Function

        Function MatchByNew(s As String) As Boolean
            Dim re As RegExp
            Set re = New RegExp
            re.Pattern = "[a-z]+"
            MatchByNew = re.Test(s)
        End Function
    `;
    assert.strictEqual(
        runFunc(code, 'MatchByCreateObject', ['hello']),
        runFunc(code, 'MatchByNew', ['hello']),
        'CreateObject と New RegExp は同じ結果'
    );
    console.log('[PASS] CreateObject と New が同じ動作');
}

console.log('\n✅ VBScript.RegExp Mock: 全テスト通過');
