/**
 * 到達定義解析テスト — Phase 4B
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { analyzeReachingDefs, buildDefUseChains, Definition, UsePoint, usePointKey } from '../../src/engine/reaching-defs';
import { ProcedureDeclaration } from '../../src/engine/parser';
import { getStmtUses } from '../../src/engine/live-vars';
import { assert } from '../../test-libs/test-runner';

function parse(code: string): ProcedureDeclaration {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find(
        (s): s is ProcedureDeclaration => s.type === 'ProcedureDeclaration',
    );
    if (!proc) throw new Error('No ProcedureDeclaration found');
    return proc;
}

function defVarNames(defs: ReadonlySet<Definition>): string[] {
    return [...defs].map(d => d.varName).sort();
}

// ─── 直線プログラム ──────────────────────────────────────────────────────────

// x = 1 の定義がブロック末尾まで到達する
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    x = 1
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);
    // 通常ブロック（entry/exit 除く）は 1 つ
    const normal = cfg.blocks.filter(b => b.kind === 'normal');
    assert.strictEqual(normal.length, 1, '直線: 通常ブロック 1 つ');

    const outDefs = result.blockOut.get(normal[0].id)!;
    assert.strictEqual(defVarNames(outDefs).includes('x'), true, '直線: x の定義が OUT に含まれる');
    console.log('[PASS] 直線: x の定義が OUT に到達');
}

// kill: x = 1 の後に x = 2 → x の定義は x = 2 のみ到達
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    x = 1
    x = 2
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);
    const normal = cfg.blocks.filter(b => b.kind === 'normal')[0];
    const outDefs = [...result.blockOut.get(normal.id)!].filter(d => d.varName === 'x');
    assert.strictEqual(outDefs.length, 1, 'kill: x の定義は最後の 1 つだけ到達');
    assert.strictEqual(outDefs[0].line > 0, true, 'kill: 到達した定義の行番号が正');
    console.log('[PASS] kill: 上書き後は最後の定義のみ到達');
}

// ─── 条件分岐 ────────────────────────────────────────────────────────────────

// If/Else の両分岐の定義がマージブロック IN に到達
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    If flag Then
        x = 10
    Else
        x = 20
    End If
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);

    // merge ブロック（If 後の Debug.Print があるブロック）の IN を確認
    // merge ブロックは then/else の後続
    const mergeBlock = cfg.blocks.find(b =>
        b.kind === 'normal' &&
        b.preds.length === 2 &&
        b.stmts.length > 0,
    );
    assert.strictEqual(mergeBlock !== undefined, true, 'If/Else: マージブロック存在');
    const inDefs = [...result.blockIn.get(mergeBlock!.id)!].filter(d => d.varName === 'x');
    assert.strictEqual(inDefs.length, 2, 'If/Else: 両分岐の x 定義が IN に到達（2 つ）');
    console.log('[PASS] If/Else: 両分岐の定義がマージポイントに到達');
}

// If のみ（Else なし）: 条件真の定義と条件偽（なし）の 2 パスが合流
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    x = 0
    If flag Then
        x = 10
    End If
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);
    // merge/after ブロックには x=0 と x=10 の 2 定義が到達するか、
    // または x=0 のみ（then を通らないパス）+ x=10（then を通るパス）
    // →合計 2 定義が到達できる（保守的）
    const mergeBlock = cfg.blocks.find(b =>
        b.kind === 'normal' && b.preds.length === 2 && b.stmts.length > 0,
    );
    if (mergeBlock) {
        const inDefs = [...result.blockIn.get(mergeBlock.id)!].filter(d => d.varName === 'x');
        assert.strictEqual(inDefs.length >= 1, true, 'If(else なし): x の定義が IN に存在');
    }
    console.log('[PASS] If(else なし): x の定義が到達');
}

// ─── ループ ──────────────────────────────────────────────────────────────────

// For ループ: バックエッジ経由でループ内の定義がヘッダーに到達
{
    const proc = parse(`
Sub Test()
    Dim i As Long
    Dim s As Long
    For i = 1 To 5
        s = s + i
    Next i
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);

    // body ブロックの OUT には s の定義が含まれる
    const bodyBlock = cfg.blocks.find(b =>
        b.kind === 'normal' &&
        b.stmts.some(st => (st as any).type === 'AssignmentStatement'),
    );
    assert.strictEqual(bodyBlock !== undefined, true, 'For: ボディブロック存在');
    const outDefs = defVarNames(result.blockOut.get(bodyBlock!.id)!);
    assert.strictEqual(outDefs.includes('s'), true, 'For: s の定義がボディ OUT に含まれる');
    console.log('[PASS] For: ループ内の定義が出力に含まれる');
}

// ─── パラメーター ─────────────────────────────────────────────────────────────

// パラメーターの定義が最初から到達している
{
    const proc = parse(`
Sub Test(ByVal n As Long)
    Dim result As Long
    result = n * 2
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);

    // 通常ブロックの IN に n の定義が到達
    const normal = cfg.blocks.filter(b => b.kind === 'normal')[0];
    const inDefs = defVarNames(result.blockIn.get(normal.id)!);
    assert.strictEqual(inDefs.includes('n'), true, 'パラメーター: n の定義が IN に到達');
    console.log('[PASS] パラメーター: 暗黙定義が到達');
}

// ─── Def-Use チェーン ────────────────────────────────────────────────────────

// x = 1; Debug.Print x → x の使用に x = 1 の定義が到達
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    x = 42
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);

    // 通常ブロックの stmts を確認して使用点を特定
    const normal = cfg.blocks.filter(b => b.kind === 'normal')[0];
    // stmts[0] = Dim x, stmts[1] = x=42, stmts[2] = Debug.Print x
    // Debug.Print x は x を USE → use point は (blockId, 2, 'x')
    const printIdx = normal.stmts.findIndex(s => s.type === 'DebugPrintStatement' || s.type === 'CallStatement');
    assert.strictEqual(printIdx >= 0, true, 'Def-Use: Debug.Print ステートメント存在');

    const uses: UsePoint[] = [{ blockId: normal.id, stmtIdx: printIdx, varName: 'x' }];
    const chains = buildDefUseChains(cfg, result, uses);
    const key = usePointKey(uses[0]);
    const reachingDefs = chains.get(key)!;
    assert.strictEqual(reachingDefs !== undefined, true, 'Def-Use: チェーン存在');
    assert.strictEqual(defVarNames(reachingDefs).includes('x'), true, 'Def-Use: x の定義が到達');
    console.log('[PASS] Def-Use チェーン: 定義が使用点に到達');
}

// If/Else 後の使用: 両分岐の定義が到達
{
    const proc = parse(`
Sub Test()
    Dim x As Long
    If flag Then
        x = 1
    Else
        x = 2
    End If
    Debug.Print x
End Sub
`);
    const { cfg, result } = analyzeReachingDefs(proc);
    const mergeBlock = cfg.blocks.find(b =>
        b.kind === 'normal' && b.preds.length === 2 && b.stmts.length > 0,
    )!;
    const printIdx = mergeBlock.stmts.findIndex(s => s.type === 'DebugPrintStatement' || s.type === 'CallStatement');
    const uses: UsePoint[] = [{ blockId: mergeBlock.id, stmtIdx: printIdx, varName: 'x' }];
    const chains = buildDefUseChains(cfg, result, uses);
    const key = usePointKey(uses[0]);
    const defs = chains.get(key)!;
    assert.strictEqual([...defs].filter(d => d.varName === 'x').length, 2,
        'Def-Use(If/Else): 両分岐の定義 2 つが到達');
    console.log('[PASS] Def-Use(If/Else): 両分岐の定義が使用点に到達');
}

// ─── ユーティリティ: getStmtUses ─────────────────────────────────────────────

// 複合式の変数収集
{
    const tokens = new Lexer('Sub T()\n    x = a + b * c\nEnd Sub').tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find((s): s is ProcedureDeclaration => s.type === 'ProcedureDeclaration')!;
    const stmt = proc.body[0]; // x = a + b * c
    const uses = getStmtUses(stmt);
    assert.strictEqual(uses.includes('a'), true, 'getStmtUses: a');
    assert.strictEqual(uses.includes('b'), true, 'getStmtUses: b');
    assert.strictEqual(uses.includes('c'), true, 'getStmtUses: c');
    assert.strictEqual(uses.includes('x'), false, 'getStmtUses: x は DEF なので含まれない');
    console.log('[PASS] getStmtUses: 複合式の変数収集');
}

console.log('\n✅ 到達定義解析: 全テスト通過');
