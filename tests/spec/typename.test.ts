import { vbaMissing } from '../../src/engine/evaluator';
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// TypeName のテスト
{
    const code = `
        Function TestBasic()
            Debug.Print TypeName(10)
            Debug.Print TypeName(10.5)
            Debug.Print TypeName("hello")
            Debug.Print TypeName(True)
        End Function
        
        Function TestNull()
            TestNull = TypeName(Null)
        End Function
        
        Function TestNothing()
            TestNothing = TypeName(Nothing)
        End Function
        
        Function TestArray()
            Dim arr(5)
            TestArray = TypeName(arr)
        End Function

        Function TestMissing(Optional x)
            TestMissing = TypeName(x)
        End Function
    `;
    
    const ev = evalVBA(code);
    // Basic types
    assert.strictEqual(ev.evalExpression('TypeName(10)'), 'Integer', 'TypeName(10) -> Integer');
    assert.strictEqual(ev.evalExpression('TypeName("abc")'), 'String', 'TypeName("abc") -> String');
    assert.strictEqual(ev.evalExpression('TypeName(True)'), 'Boolean', 'TypeName(True) -> Boolean');

    // Null/Nothing/Empty
    assert.strictEqual(runFunc(code, 'TestNull'), 'Null', 'TypeName(Null) -> Null');
    assert.strictEqual(runFunc(code, 'TestNothing'), 'Nothing', 'TypeName(Nothing) -> Nothing');
    
    // Array
    assert.strictEqual(runFunc(code, 'TestArray'), 'Variant()', 'TypeName(Array) -> Variant()');

    // Missing
    assert.strictEqual(runFunc(code, 'TestMissing', []), 'Error', 'TypeName(Missing) -> Error');

    console.log('[PASS] TypeName の基本動作');
}

// 数値リテラルのサフィックス型情報保持
{
    const ev = evalVBASingle('Function Dummy(): End Function');
    assert.strictEqual(ev.evalExpression('TypeName(100%)'), 'Integer', 'TypeName(100%) -> Integer');
    assert.strictEqual(ev.evalExpression('TypeName(100&)'), 'Long', 'TypeName(100&) -> Long');
    assert.strictEqual(ev.evalExpression('TypeName(1.5!)'), 'Single', 'TypeName(1.5!) -> Single');
    assert.strictEqual(ev.evalExpression('TypeName(1.5#)'), 'Double', 'TypeName(1.5#) -> Double');
    assert.strictEqual(ev.evalExpression('TypeName(1.5@)'), 'Currency', 'TypeName(1.5@) -> Currency');
    assert.strictEqual(ev.evalExpression('TypeName(123^)'), 'LongLong', 'TypeName(123^) -> LongLong');

    // VarType も確認
    assert.strictEqual(ev.evalExpression('VarType(100%)'), 2, 'VarType(100%) -> 2 (vbInteger)');
    assert.strictEqual(ev.evalExpression('VarType(100&)'), 3, 'VarType(100&) -> 3 (vbLong)');
    assert.strictEqual(ev.evalExpression('VarType(1.5!)'), 4, 'VarType(1.5!) -> 4 (vbSingle)');
    assert.strictEqual(ev.evalExpression('VarType(123^)'), 20, 'VarType(123^) -> 20 (vbLongLong)');

    // 小数点付きリテラルは isFloat により Double
    assert.strictEqual(ev.evalExpression('TypeName(1.0)'), 'Double', 'TypeName(1.0) -> Double');
    assert.strictEqual(ev.evalExpression('TypeName(1.5)'), 'Double', 'TypeName(1.5) -> Double');

    console.log('[PASS] 数値リテラルのサフィックス型情報保持');
}

// サフィックス付きリテラルのオーバーフロー検出
{
    function overflows(code: string): boolean {
        try {
            evalVBASingle(code).callProcedure('F', []);
            return false;
        } catch (e: any) {
            return /Overflow/.test(e.message);
        }
    }
    assert.strictEqual(overflows('Function F(): F = 100000%: End Function'), true, '100000% -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 40000%: End Function'), true, '40000% (>32767) -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 3000000000&: End Function'), true, '3000000000& -> Overflow');
    assert.strictEqual(overflows('Function F(): F = 100&: End Function'), false, '100& -> no overflow');
    assert.strictEqual(overflows('Function F(): F = 32767%: End Function'), false, '32767% -> no overflow');

    console.log('[PASS] サフィックス付きリテラルのオーバーフロー検出');
}

// Bug 22-1 レグレッション: Variant 変数へのサブタイプ追跡（代入時点の RHS 型を保持）
{
    const code = `
        Function GetIntLiteral() As String
            Dim v As Variant
            v = 42
            GetIntLiteral = TypeName(v)
        End Function

        Function GetVarType42() As Integer
            Dim v As Variant
            v = 42
            GetVarType42 = VarType(v)
        End Function

        Function GetLongLiteral() As String
            Dim v As Variant
            v = 40000
            GetLongLiteral = TypeName(v)
        End Function

        Function GetDoubleLiteral() As String
            Dim v As Variant
            v = 3.14
            GetDoubleLiteral = TypeName(v)
        End Function

        Function GetDoubleDiv() As String
            ' 6.0 / 2.0 は Double 除算 → JS では 3 (Number.isInteger=true) だが型は Double であるべき
            Dim v As Variant
            v = 6.0 / 2.0
            GetDoubleDiv = TypeName(v)
        End Function

        Function GetCLngResult() As String
            Dim v As Variant
            v = CLng(42.9)
            GetCLngResult = TypeName(v)
        End Function

        Function GetPropagated() As String
            ' Variant から Variant への代入でサブタイプが伝播する
            Dim src As Variant, dst As Variant
            src = 42
            dst = src
            GetPropagated = TypeName(dst)
        End Function
    `;

    const ev = evalVBASingle(code);
    // Integer リテラル → "Integer"
    assert.strictEqual(ev.callProcedure('GetIntLiteral', []), 'Integer',
        'v = 42: TypeName(v) -> Integer');
    assert.strictEqual(ev.callProcedure('GetVarType42', []), 2,
        'v = 42: VarType(v) -> 2 (vbInteger)');
    // Integer 範囲外の整数リテラル → "Long"
    assert.strictEqual(ev.callProcedure('GetLongLiteral', []), 'Long',
        'v = 40000: TypeName(v) -> Long');
    // Double リテラル → "Double"
    assert.strictEqual(ev.callProcedure('GetDoubleLiteral', []), 'Double',
        'v = 3.14: TypeName(v) -> Double');
    // Double 演算の結果（JS では整数値だが型は Double） → "Double"
    assert.strictEqual(ev.callProcedure('GetDoubleDiv', []), 'Double',
        'v = 6.0/2.0: TypeName(v) -> Double (not Integer)');
    // CLng() の戻り型 → "Long"
    assert.strictEqual(ev.callProcedure('GetCLngResult', []), 'Long',
        'v = CLng(42.9): TypeName(v) -> Long');
    // Variant→Variant 代入でサブタイプが伝播 → "Integer"
    assert.strictEqual(ev.callProcedure('GetPropagated', []), 'Integer',
        'dst = src (src=42): TypeName(dst) -> Integer');

    console.log('[PASS] Bug 22-1: Variant 変数の TypeName/VarType サブタイプ追跡（代入時点の型を保持）');
}

// BinaryExpression の VBA 型昇格規則
{
    const ev = evalVBASingle(`
        Function T(v As Variant) As String: T = TypeName(v): End Function
        Function FSlash():    Dim v As Variant: v = 6 / 2:        FSlash    = T(v): End Function
        Function FBSlash():   Dim v As Variant: v = 6 \\ 2:       FBSlash   = T(v): End Function
        Function FAdd():      Dim v As Variant: v = 3 + 4:        FAdd      = T(v): End Function
        Function FSub():      Dim v As Variant: v = 5 - 2:        FSub      = T(v): End Function
        Function FMul():      Dim v As Variant: v = 3 * 4:        FMul      = T(v): End Function
        Function FMod():      Dim v As Variant: v = 7 Mod 3:      FMod      = T(v): End Function
        Function FPow():      Dim v As Variant: v = 2 ^ 3:        FPow      = T(v): End Function
        Function FLongAdd():  Dim v As Variant: v = 40000 + 1:    FLongAdd  = T(v): End Function
        Function FDblDiv():   Dim v As Variant: v = 6.0 / 2.0:    FDblDiv   = T(v): End Function
        Function FSngDiv():   Dim v As Variant: v = CSng(6) / CSng(2): FSngDiv = T(v): End Function
        Function FNeg():      Dim v As Variant: v = -42:           FNeg      = T(v): End Function
    `);
    assert.strictEqual(ev.callProcedure('FSlash',   []), 'Double',  '6 / 2   -> Double (/ は常に Double)');
    assert.strictEqual(ev.callProcedure('FBSlash',  []), 'Integer', '6 \\ 2  -> Integer (整数除算)');
    assert.strictEqual(ev.callProcedure('FAdd',     []), 'Integer', '3 + 4   -> Integer');
    assert.strictEqual(ev.callProcedure('FSub',     []), 'Integer', '5 - 2   -> Integer');
    assert.strictEqual(ev.callProcedure('FMul',     []), 'Integer', '3 * 4   -> Integer');
    assert.strictEqual(ev.callProcedure('FMod',     []), 'Integer', '7 Mod 3 -> Integer');
    assert.strictEqual(ev.callProcedure('FPow',     []), 'Double',  '2 ^ 3   -> Double (^ は常に Double)');
    assert.strictEqual(ev.callProcedure('FLongAdd', []), 'Long',    '40000+1 -> Long');
    assert.strictEqual(ev.callProcedure('FDblDiv',  []), 'Double',  '6.0/2.0 -> Double');
    assert.strictEqual(ev.callProcedure('FSngDiv',  []), 'Single',  'CSng/CSng -> Single');
    assert.strictEqual(ev.callProcedure('FNeg',     []), 'Integer', '-42     -> Integer (単項マイナス)');
    console.log('[PASS] BinaryExpression の VBA 型昇格規則');
}

// Bug CG: TypeName/VarType on Variant array elements lost the numeric subtype (always returned Double).
// Assignment to arr(i) = 42 now tracks the subtype in arr.__vbaSubtypes__[key].
{
    const code = `
Function TestCG_DirectLiterals() As String
    Dim arr(4) As Variant
    arr(0) = 42
    arr(1) = 100000
    arr(2) = 3.14
    arr(3) = True
    arr(4) = "hello"
    TestCG_DirectLiterals = TypeName(arr(0)) & "," & TypeName(arr(1)) & "," & TypeName(arr(2)) & "," & TypeName(arr(3)) & "," & TypeName(arr(4))
End Function
Function TestCG_VarType() As String
    Dim arr(1) As Variant
    arr(0) = 42
    arr(1) = 40000
    TestCG_VarType = VarType(arr(0)) & "," & VarType(arr(1))
End Function
`;
    assert.strictEqual(
        evalVBASingle(code).callProcedure('TestCG_DirectLiterals', []),
        'Integer,Long,Double,Boolean,String',
        'Bug CG: TypeName on Variant array elements preserves numeric subtype'
    );
    assert.strictEqual(
        evalVBASingle(code).callProcedure('TestCG_VarType', []),
        '2,3',
        'Bug CG: VarType on Variant array elements: Integer=2, Long=3'
    );
    console.log('[PASS] Bug CG: Variant 配列要素の TypeName/VarType が正しいサブタイプを返す');
}

// Bug 30-A: VarType(classInstance) が 36 (vbUserDefinedType) を返す
// Class インスタンスは vbObject = 9 を返すべき
{
    const clsCode = `Public Value As Long\n`;
    const mainCode = `
Function GetVarType() As Long
    Dim obj As New MyClass30A
    GetVarType = VarType(obj)
End Function
Function GetTypeName() As String
    Dim obj As New MyClass30A
    GetTypeName = TypeName(obj)
End Function
`;
    const ev = evalVBAModules([
        { name: 'MyClass30A', code: clsCode, parseAsClass: 'MyClass30A' },
        { name: 'Main', code: mainCode },
    ]);
    assert.strictEqual(ev.callProcedure('GetVarType', []), 9, 'Bug 30-A: VarType(classInstance) = 9 (vbObject)');
    assert.strictEqual(ev.callProcedure('GetTypeName', []), 'MyClass30A', 'Bug 30-A: TypeName(classInstance) = class name');
    console.log('[PASS] Bug 30-A: VarType(class instance) = 9 (vbObject)');
}

console.log('\n✅ TypeName: 全テスト通過');
