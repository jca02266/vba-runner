/**
 * FlowEnvironment (src/engine/flow-environment.ts) のテスト
 * Phase 2: 代入フロー追跡による型推定
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { buildTypeEnvironment } from '../../src/engine/type-environment';
import {
    buildFlowEnvironment,
    inferExprType,
    resolveVarType,
} from '../../src/engine/flow-environment';
import { assert } from '../../test-libs/test-runner';

function buildEnvs(code: string, procName: string) {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const typeEnv = buildTypeEnvironment(ast);
    const proc = (ast.body as any[]).find(
        s => s.type === 'ProcedureDeclaration' && s.name.name === procName
    );
    if (!proc) throw new Error(`Procedure ${procName} not found`);
    const flowEnv = buildFlowEnvironment(proc, typeEnv);
    return { typeEnv, flowEnv };
}

// ─── 1. 宣言型の初期投入 ──────────────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim rng As Range
            Dim x As Long
            Dim v As Variant
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('rng')?.inferredType, 'Range', 'Range 宣言型が初期値');
    assert.strictEqual(flowEnv.get('rng')?.source, 'declaration', 'source = declaration');
    assert.strictEqual(flowEnv.get('x')?.inferredType, 'Long', 'Long 宣言型が初期値');
    assert.strictEqual(flowEnv.get('v'), undefined, 'Variant は投入されない');
    console.log('[PASS] 宣言型の初期投入');
}

// ─── 2. Set 代入による型の更新 ────────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim ws As Variant
            Dim rng As Variant
            Set ws = ActiveWorkbook.Worksheets(1)
            Set rng = ws.Range("A1")
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('ws')?.inferredType, 'Worksheet', 'Worksheets → Worksheet');
    assert.strictEqual(flowEnv.get('ws')?.source, 'assignment', 'source = assignment');
    assert.strictEqual(flowEnv.get('rng')?.inferredType, 'Range', 'ws.Range → Range');
    console.log('[PASS] Set 代入による型の更新');
}

// ─── 3. 通常代入による型の更新 ────────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim n As Variant
            Dim s As Variant
            n = 42
            s = "hello"
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('n')?.inferredType, 'Long', '整数リテラル → Long');
    assert.strictEqual(flowEnv.get('s')?.inferredType, 'String', '文字列リテラル → String');
    console.log('[PASS] 通常代入による型の更新');
}

// ─── 4. MemberExpression の既知プロパティ推定 ────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim r As Variant
            Dim c As Variant
            Set r = ActiveSheet.Cells
            Set c = ActiveSheet.UsedRange
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('r')?.inferredType, 'Range', 'Cells → Range');
    assert.strictEqual(flowEnv.get('c')?.inferredType, 'Range', 'UsedRange → Range');
    console.log('[PASS] MemberExpression 既知プロパティ');
}

// ─── 5. NewExpression による型推定 ────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim d As Variant
            Set d = New Dictionary
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('d')?.inferredType, 'Dictionary', 'New Dictionary → Dictionary');
    console.log('[PASS] NewExpression 型推定');
}

// ─── 6. CreateObject による型推定 ─────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim fso As Variant
            Set fso = CreateObject("Scripting.FileSystemObject")
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('fso')?.inferredType, 'FileSystemObject', 'CreateObject FSO');
    console.log('[PASS] CreateObject 型推定');
}

// ─── 7. コピー伝播（変数から変数への代入） ───────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim rng As Range
            Dim copy As Variant
            Set copy = rng
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('copy')?.inferredType, 'Range', '変数コピー伝播');
    console.log('[PASS] コピー伝播');
}

// ─── 8. If / For ブロック内の代入も追跡 ─────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim x As Variant
            If True Then
                x = 1
            End If
            Dim y As Variant
            For i = 1 To 10
                y = "str"
            Next i
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('x')?.inferredType, 'Long', 'If 内の代入を追跡');
    assert.strictEqual(flowEnv.get('y')?.inferredType, 'String', 'For 内の代入を追跡');
    console.log('[PASS] ブロック内代入追跡');
}

// ─── 9. resolveVarType: FlowEnv 優先 ─────────────────────────────────────────
{
    const code = `
        Sub Test()
            Dim rng As Variant
            Set rng = ActiveSheet.Cells
        End Sub
    `;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const typeEnv = buildTypeEnvironment(ast);
    const proc = (ast.body as any[])[0];
    const flowEnv = buildFlowEnvironment(proc, typeEnv);

    const t = resolveVarType('rng', typeEnv, flowEnv, 'Test');
    assert.strictEqual(t, 'Range', 'resolveVarType FlowEnv 優先');
    console.log('[PASS] resolveVarType FlowEnv 優先');
}

// ─── 10. 未知変数は Variant ──────────────────────────────────────────────────
{
    const code = `
        Sub Test()
        End Sub
    `;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const typeEnv = buildTypeEnvironment(ast);
    const proc = (ast.body as any[])[0];
    const flowEnv = buildFlowEnvironment(proc, typeEnv);

    const t = resolveVarType('undefined_var', typeEnv, flowEnv, 'Test');
    assert.strictEqual(t, 'Variant', '未知変数 → Variant');
    console.log('[PASS] 未知変数は Variant');
}

// ─── 11. パラメーター型の初期投入 ────────────────────────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test(rng As Range, n As Long)
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('rng')?.inferredType, 'Range', 'Range パラメーター');
    assert.strictEqual(flowEnv.get('n')?.inferredType, 'Long', 'Long パラメーター');
    console.log('[PASS] パラメーター型の初期投入');
}

// ─── 12. KNOWN_PROPERTY_RETURN_TYPES の主要エントリ ─────────────────────────
{
    const { flowEnv } = buildEnvs(`
        Sub Test()
            Dim r1 As Variant
            Dim r2 As Variant
            Dim r3 As Variant
            Set r1 = something.Offset(1, 0)
            Set r2 = something.EntireRow
            Set r3 = something.CurrentRegion
        End Sub
    `, 'Test');

    assert.strictEqual(flowEnv.get('r1')?.inferredType, 'Range', 'Offset → Range');
    assert.strictEqual(flowEnv.get('r2')?.inferredType, 'Range', 'EntireRow → Range');
    assert.strictEqual(flowEnv.get('r3')?.inferredType, 'Range', 'CurrentRegion → Range');
    console.log('[PASS] KNOWN_PROPERTY_RETURN_TYPES 主要エントリ');
}

console.log('\n✅ FlowEnvironment: 全テスト通過');
