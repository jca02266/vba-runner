/**
 * Unicode 識別子 (MS-VBAL §3.3.5) のテスト
 * 日本語（ひらがな・カタカナ・漢字）の変数名・関数名・UDT 名をサポートする
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// --- 1. 漢字変数名 ---
const kanjiVarCode = `
    Public 氏名 As String
    Sub Test()
        氏名 = "山田太郎"
    End Sub
`;
const ev1 = evalVBA(kanjiVarCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('氏名'), '山田太郎', '漢字変数に代入できる');
console.log('[PASS] 漢字変数名');

// --- 2. ひらがな変数名 ---
const hiraganaCode = `
    Sub Test()
        Dim ごうけい As Long
        ごうけい = 100
    End Sub
`;
const ev2 = evalVBA(hiraganaCode);
ev2.callProcedure('Test', []);
console.log('[PASS] ひらがな変数名');

// --- 3. カタカナ変数名 ---
const katakanaCode = `
    Public カウンター As Long
    Sub Test()
        Dim i As Long
        For i = 1 To 3
            カウンター = カウンター + 1
        Next i
    End Sub
`;
const ev3 = evalVBA(katakanaCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('カウンター'), 3, 'カタカナ変数がループ内で正しくインクリメントされる');
console.log('[PASS] カタカナ変数名');

// --- 4. 日本語 Function 名 ---
const japaneseFuncCode = `
    Function 合計(a As Long, b As Long) As Long
        合計 = a + b
    End Function
    Public 結果 As Long
    Sub Test()
        結果 = 合計(10, 20)
    End Sub
`;
const ev4 = evalVBA(japaneseFuncCode);
ev4.callProcedure('Test', []);
assert.strictEqual(ev4.env.get('結果'), 30, '日本語 Function を呼び出せる');
console.log('[PASS] 日本語 Function 名');

// --- 5. 日本語 Sub 名 ---
const japaneseSubCode = `
    Public 実行結果 As String
    Sub 初期化()
        実行結果 = "初期化完了"
    End Sub
    Sub Test()
        Call 初期化()
    End Sub
`;
const ev5 = evalVBA(japaneseSubCode);
ev5.callProcedure('Test', []);
assert.strictEqual(ev5.env.get('実行結果'), '初期化完了', '日本語 Sub を呼び出せる');
console.log('[PASS] 日本語 Sub 名');

// --- 6. 日本語 UDT ---
const japaneseUDTCode = `
    Type 顧客情報
        名前 As String
        年齢 As Integer
    End Type
    Public 顧客 As 顧客情報
    Sub Test()
        顧客.名前 = "鈴木花子"
        顧客.年齢 = 25
    End Sub
`;
const ev6 = evalVBA(japaneseUDTCode);
ev6.callProcedure('Test', []);
const 顧客 = ev6.env.get('顧客');
assert.strictEqual(顧客['名前'], '鈴木花子', '日本語 UDT メンバー 名前');
assert.strictEqual(顧客['年齢'], 25, '日本語 UDT メンバー 年齢');
console.log('[PASS] 日本語 UDT');

// --- 7. 日本語 + ASCII 混在識別子 ---
const mixedCode = `
    Sub Test()
        Dim 合計額 As Long
        Dim item1 As Long
        item1 = 500
        合計額 = item1 * 2
    End Sub
    Public 合計額 As Long
    Sub TestOuter()
        Dim 合計額 As Long
        Dim item1 As Long
        item1 = 500
        合計額 = item1 * 2
    End Sub
`;
const ev7 = evalVBA(mixedCode);
ev7.callProcedure('TestOuter', []);
console.log('[PASS] 日本語 + ASCII 混在識別子');

console.log('\n✅ Unicode 識別子: 全テスト通過');
