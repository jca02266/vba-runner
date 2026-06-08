/**
 * Call Statement (§5.4.2.1) & Named Parameters (§5.6.3.1.2) のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. Call Statement の基本動作 ---
const callCode = `
    Dim globalVal
    Sub SetVal(v)
        globalVal = v
    End Sub

    Sub Main()
        Call SetVal(100)
    End Sub
`;
const ev1 = evalVBA(callCode);
ev1.callProcedure('Main', []);
assert.strictEqual(ev1.env.get('globalval'), 100, 'Call SetVal(100) で変数が更新される');
console.log('[PASS] Call Statement 基本動作');

// --- 2. Named Parameters (名前付き引数) ---
const namedParamCode = `
    Function Subtract(a, b)
        Subtract = a - b
    End Function

    Function TestNamed()
        TestNamed = Subtract(b:=10, a:=50)
    End Function
`;
assert.strictEqual(runFunc(namedParamCode, 'TestNamed'), 40, 'b:=10, a:=50 で 50-10=40');
console.log('[PASS] Named Parameters 基本動作');

// --- 3. 名前付き引数と通常引数の混在 ---
const mixedCode = `
    Function Mixed(a, b, Optional c = 0)
        Mixed = a + b + c
    End Function

    Function TestMixed()
        TestMixed = Mixed(10, c:=30, b:=20)
    End Function
`;
assert.strictEqual(runFunc(mixedCode, 'TestMixed'), 60, '10 + 20 + 30 = 60');
console.log('[PASS] 通常引数と名前付き引数の混在');

// --- 4. Call と Named Params の組み合わせ ---
const callNamedCode = `
    Dim result
    Sub Store(val)
        result = val
    End Sub

    Sub Main()
        Call Store(val:=999)
    End Sub
`;
const ev4 = evalVBA(callNamedCode);
ev4.callProcedure('Main', []);
assert.strictEqual(ev4.env.get('result'), 999, 'Call Store(val:=999) が動作する');
console.log('[PASS] Call + Named Parameters');

// --- 5. 名前付き引数名がキーワードの場合（COM/組み込みメソッド用途） ---
// MS-VBAL §5.6.13.1: named-argument ::= unrestricted-name ':=' expression
// unrestricted-name はキーワードを含む。Type・Date・Name など VBA キーワードが
// パラメーター名になる COM メソッド（Validation.Add 等）の呼び出しに必要。
// ユーザー定義関数でキーワードをパラメーター名にすることはできないが、
// パーサーは呼び出し側で unrestricted-name を許容しなければならない。
{
    // With ブロック内の .Add Type:=... — 実際の用途: With Selection.Validation : .Add Type:=xlValidateList
    const src = `
Sub Test()
    With Selection.Validation
        .Add Type:=1, AlertStyle:=2, Formula1:="A,B"
    End With
End Sub
`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    assert.strictEqual(ast.diagnostics?.length ?? 0, 0, 'Type:= を含む名前付き引数がパースエラーにならない');

    const proc = ast.body[0] as any;
    const withStmt = proc.body[0] as any;
    const callStmt = withStmt.body[0] as any;
    const args: any[] = callStmt.expression?.args ?? [];
    assert.strictEqual(args.length, 3, '引数が3つ認識される');
    assert.strictEqual(args[0].type, 'NamedArgument', '第1引数が NamedArgument');
    assert.strictEqual(args[0].name, 'Type', '第1引数名が Type');
    assert.strictEqual(args[1].name, 'AlertStyle', '第2引数名が AlertStyle');
    assert.strictEqual(args[2].name, 'Formula1', '第3引数名が Formula1');
    console.log('[PASS] キーワードを名前とする名前付き引数（Type:=, Date:=, Name:= 等）');
}

// --- 6. 他のキーワード（Date, Name）も名前付き引数として認識される ---
{
    const src = `Sub Test() : Call Foo(Date:=1, Name:="x") : End Sub`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    assert.strictEqual(ast.diagnostics?.length ?? 0, 0, 'Date:= / Name:= もパースエラーにならない');
    const args: any[] = (ast.body[0] as any).body[0].expression?.args ?? [];
    assert.strictEqual(args[0].name, 'Date', 'Date が引数名として認識される');
    assert.strictEqual(args[1].name, 'Name', 'Name が引数名として認識される');
    console.log('[PASS] Date:= / Name:= も名前付き引数として認識される');
}

// --- 7. 数値リテラルは名前付き引数の名前として扱われない ---
{
    const src = `Sub Test() : Call Foo(1 + 2) : End Sub`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    // 数値の後に ':=' が来ないケースで誤検知しないことを確認
    assert.strictEqual(ast.diagnostics?.length ?? 0, 0, '数値引数は通常の式として扱われる');
    const args: any[] = (ast.body[0] as any).body[0].expression?.args ?? [];
    assert.strictEqual(args[0].type, 'BinaryExpression', '1 + 2 は BinaryExpression として解析される');
    console.log('[PASS] 数値リテラルは名前付き引数名として扱われない');
}

console.log('\n✅ Call & Named Parameters: 全テスト通過');
