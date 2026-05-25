/**
 * TypeEnvironment (src/engine/type-environment.ts) のテスト
 * Phase 1: 宣言ベースの型環境構築
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { buildTypeEnvironment, lookupType, isNumericType } from '../../src/engine/type-environment';
import { assert } from '../../test-libs/test-runner';

function buildEnv(code: string) {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    return buildTypeEnvironment(ast);
}

// ─── 1. モジュールレベルの変数宣言 ───────────────────────────────────────────
const env1 = buildEnv(`
    Public count As Long
    Private name As String
    Dim ratio As Double
    Dim flag As Boolean
    Dim rng As Range
    Dim v As Variant
    Dim noType
`);

assert.strictEqual(lookupType(env1, 'count')?.declaredType, 'Long',    'Long 宣言');
assert.strictEqual(lookupType(env1, 'name')?.declaredType,  'String',  'String 宣言');
assert.strictEqual(lookupType(env1, 'ratio')?.declaredType, 'Double',  'Double 宣言');
assert.strictEqual(lookupType(env1, 'flag')?.declaredType,  'Boolean', 'Boolean 宣言');
assert.strictEqual(lookupType(env1, 'rng')?.declaredType,   'Range',   'オブジェクト型');
assert.strictEqual(lookupType(env1, 'v')?.declaredType,     'Variant', '明示 Variant');
assert.strictEqual(lookupType(env1, 'noType')?.declaredType,'Variant', '型なし → Variant');
assert.strictEqual(lookupType(env1, 'count')?.kind, 'variable', 'kind = variable');
assert.strictEqual(lookupType(env1, 'COUNT')?.declaredType, 'Long', '大文字小文字を無視');
console.log('[PASS] モジュールレベル変数宣言');

// ─── 2. Const 宣言 ────────────────────────────────────────────────────────────
const env2 = buildEnv(`
    Const LIMIT = 100
    Const RATE = 3.14
    Const LABEL = "hello"
    Const NEG = -50
    Const FLAG = True
`);

const limit = lookupType(env2, 'LIMIT');
assert.strictEqual(limit?.kind,         'const', 'LIMIT kind = const');
assert.strictEqual(limit?.constValue,   100,     'LIMIT constValue = 100');
assert.strictEqual(limit?.declaredType, 'Long',  'LIMIT 整数 → Long');

assert.strictEqual(lookupType(env2, 'RATE')?.constValue,   3.14,    'RATE constValue');
assert.strictEqual(lookupType(env2, 'RATE')?.declaredType, 'Double','RATE 小数 → Double');
assert.strictEqual(lookupType(env2, 'LABEL')?.constValue,  'hello', 'LABEL constValue');
assert.strictEqual(lookupType(env2, 'NEG')?.constValue,    -50,     'NEG 負数');
assert.strictEqual(lookupType(env2, 'FLAG')?.constValue,   true,    'FLAG = True');
console.log('[PASS] Const 宣言');

// ─── 3. Function / Sub 宣言 ───────────────────────────────────────────────────
const env3 = buildEnv(`
    Function GetTotal() As Double
    End Function
    Function GetName() As String
    End Function
    Sub DoSomething()
    End Sub
    Function NoReturn()
    End Function
`);

assert.strictEqual(lookupType(env3, 'GetTotal')?.kind,        'function', 'Function kind');
assert.strictEqual(lookupType(env3, 'GetTotal')?.declaredType,'Double',   'Function 戻り型');
assert.strictEqual(lookupType(env3, 'GetName')?.declaredType, 'String',   'String 戻り型');
assert.strictEqual(lookupType(env3, 'NoReturn')?.declaredType,'Variant',  '戻り型なし → Variant');
assert.strictEqual(lookupType(env3, 'DoSomething'),            undefined,  'Sub は登録されない');
console.log('[PASS] Function/Sub 宣言');

// ─── 4. プロシージャ内のローカル宣言 ──────────────────────────────────────────
const env4 = buildEnv(`
    Sub MyProc(count As Long, name As String)
        Dim local As Integer
        Const THRESH = 200
    End Sub
`);

// ローカルスコープで検索
assert.strictEqual(lookupType(env4, 'local',  'MyProc')?.declaredType, 'Integer', 'ローカル変数');
assert.strictEqual(lookupType(env4, 'THRESH', 'MyProc')?.constValue,   200,       'ローカル Const');
assert.strictEqual(lookupType(env4, 'count',  'MyProc')?.kind,         'parameter','パラメーター');
assert.strictEqual(lookupType(env4, 'count',  'MyProc')?.declaredType, 'Long',    'パラメーター型');
assert.strictEqual(lookupType(env4, 'name',   'MyProc')?.declaredType, 'String',  'String パラメーター');
// モジュールレベルでは見えない
assert.strictEqual(lookupType(env4, 'local'),  undefined, 'ローカル変数はモジュール非公開');
console.log('[PASS] プロシージャ内ローカル宣言');

// ─── 5. スコープの優先順位（ローカル > モジュール）────────────────────────────
const env5 = buildEnv(`
    Dim x As String
    Sub Test()
        Dim x As Long
    End Sub
`);

assert.strictEqual(lookupType(env5, 'x')?.declaredType,         'String', 'モジュールレベル');
assert.strictEqual(lookupType(env5, 'x', 'Test')?.declaredType, 'Long',   'ローカルがモジュールを隠蔽');
console.log('[PASS] スコープの優先順位');

// ─── 6. Enum 宣言 ─────────────────────────────────────────────────────────────
const env6 = buildEnv(`
    Enum Priority
        Low
        Medium
        High = 10
        Critical
    End Enum
`);

assert.strictEqual(lookupType(env6, 'Low')?.kind,           'enum-member', 'Low kind');
assert.strictEqual(lookupType(env6, 'Low')?.constValue,     0,             'Low = 0');
assert.strictEqual(lookupType(env6, 'Medium')?.constValue,  1,             'Medium = 1');
assert.strictEqual(lookupType(env6, 'High')?.constValue,    10,            'High = 10');
assert.strictEqual(lookupType(env6, 'Critical')?.constValue,11,            'Critical = 11');
assert.strictEqual(lookupType(env6, 'Low')?.declaredType,   'Long',        'Enum メンバーは Long');
console.log('[PASS] Enum 宣言');

// ─── 7. 配列宣言 ──────────────────────────────────────────────────────────────
const env7 = buildEnv(`
    Dim items(10) As String
    Dim dynamic() As Long
`);

assert.strictEqual(lookupType(env7, 'items')?.isArray, true,   '固定配列 isArray');
assert.strictEqual(lookupType(env7, 'dynamic')?.isArray, true, '動的配列 isArray');
assert.strictEqual(lookupType(env7, 'items')?.declaredType, 'String', '配列の要素型');
console.log('[PASS] 配列宣言');

// ─── 8. isNumericType ヘルパー ────────────────────────────────────────────────
assert.strictEqual(isNumericType('Long'),    true,  'Long は数値型');
assert.strictEqual(isNumericType('Double'),  true,  'Double は数値型');
assert.strictEqual(isNumericType('Integer'), true,  'Integer は数値型');
assert.strictEqual(isNumericType('String'),  false, 'String は数値型でない');
assert.strictEqual(isNumericType('Range'),   false, 'Range は数値型でない');
assert.strictEqual(isNumericType(undefined), false, 'undefined は数値型でない');
console.log('[PASS] isNumericType');

// ─── 9. ブロック内ローカル宣言（If/For の中） ─────────────────────────────────
const env9 = buildEnv(`
    Sub Test()
        Dim outer As Long
        If True Then
            Dim inner As String
        End If
        For i = 1 To 10
            Dim loopVar As Double
        Next i
    End Sub
`);

assert.strictEqual(lookupType(env9, 'outer',   'Test')?.declaredType, 'Long',   'If 外の宣言');
assert.strictEqual(lookupType(env9, 'inner',   'Test')?.declaredType, 'String', 'If 内の宣言');
assert.strictEqual(lookupType(env9, 'loopVar', 'Test')?.declaredType, 'Double', 'For 内の宣言');
console.log('[PASS] ブロック内ローカル宣言');

console.log('\n✅ TypeEnvironment: 全テスト通過');
