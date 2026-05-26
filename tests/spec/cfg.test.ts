/**
 * CFG（制御フローグラフ）構築テスト — Phase 4A
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { buildCFG, BasicBlock, CFG } from '../../src/engine/cfg';
import { ProcedureDeclaration } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function buildFromSource(code: string): CFG {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find(
        (s): s is ProcedureDeclaration =>
            s.type === 'ProcedureDeclaration',
    );
    if (!proc) throw new Error('No ProcedureDeclaration found');
    return buildCFG(proc);
}

function reachable(cfg: CFG): Set<number> {
    const visited = new Set<number>();
    const queue: BasicBlock[] = [cfg.entry];
    while (queue.length) {
        const b = queue.pop()!;
        if (visited.has(b.id)) continue;
        visited.add(b.id);
        for (const s of b.succs) queue.push(s);
    }
    return visited;
}

// ─── 直線プログラム ──────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Dim x As Long
    x = 1
    x = 2
End Sub
`);
    // entry → first → exit
    assert.strictEqual(cfg.blocks.length, 3, '直線: ブロック数 3');
    assert.strictEqual(cfg.entry.succs.length, 1, '直線: entry の後継 1 つ');
    assert.strictEqual(cfg.exit.preds.length, 1, '直線: exit の先行 1 つ');
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, '直線: exit に到達可能');
    console.log('[PASS] 直線プログラム');
}

// ─── If (Else なし) ──────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    If x > 0 Then
        x = 1
    End If
    x = 2
End Sub
`);
    // entry, first(if条件), then, merge(x=2), exit
    assert.strictEqual(cfg.blocks.length >= 4, true, 'If(else なし): ブロック数 >= 4');
    // 条件ブロックから thenBlock と mergeBlock の 2 つへのエッジ
    const condBlock = cfg.entry.succs[0];
    assert.strictEqual(condBlock.succs.length, 2, 'If(else なし): 条件ブロックの後継 2');
    console.log('[PASS] If(Else なし)');
}

// ─── If...Else ───────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    If x > 0 Then
        x = 1
    Else
        x = 2
    End If
    x = 3
End Sub
`);
    // entry, first(if条件), then, else, merge, exit
    assert.strictEqual(cfg.blocks.length >= 5, true, 'If/Else: ブロック数 >= 5');
    const condBlock = cfg.entry.succs[0];
    assert.strictEqual(condBlock.succs.length, 2, 'If/Else: 条件ブロックの後継 2');
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'If/Else: exit に到達可能');
    console.log('[PASS] If...Else');
}

// ─── For...Next ──────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Dim i As Long
    For i = 1 To 10
        i = i + 1
    Next i
    Dim y As Long
End Sub
`);
    // entry, first(Dim), header(For), body, afterBlock, (Dim y might be in afterBlock), exit
    assert.strictEqual(cfg.blocks.length >= 4, true, 'For: ブロック数 >= 4');

    // ヘッダーブロックにバックエッジがあること
    let hasBackEdge = false;
    for (const b of cfg.blocks) {
        for (const s of b.succs) {
            if (s.id < b.id) { hasBackEdge = true; break; }
        }
    }
    assert.strictEqual(hasBackEdge, true, 'For: バックエッジ存在');
    console.log('[PASS] For...Next');
}

// ─── Exit For ────────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Dim i As Long
    For i = 1 To 10
        If i = 5 Then Exit For
    Next i
End Sub
`);
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'Exit For: exit に到達可能');
    console.log('[PASS] Exit For');
}

// ─── Do While...Loop ─────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Do While x < 10
        x = x + 1
    Loop
End Sub
`);
    assert.strictEqual(cfg.blocks.length >= 4, true, 'DoWhile: ブロック数 >= 4');

    let hasBackEdge = false;
    for (const b of cfg.blocks) {
        for (const s of b.succs) {
            if (s.id < b.id) { hasBackEdge = true; break; }
        }
    }
    assert.strictEqual(hasBackEdge, true, 'DoWhile: バックエッジ存在');
    console.log('[PASS] Do While...Loop');
}

// ─── While...Wend ────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    While x < 10
        x = x + 1
    Wend
End Sub
`);
    let hasBackEdge = false;
    for (const b of cfg.blocks) {
        for (const s of b.succs) {
            if (s.id < b.id) { hasBackEdge = true; break; }
        }
    }
    assert.strictEqual(hasBackEdge, true, 'While: バックエッジ存在');
    console.log('[PASS] While...Wend');
}

// ─── Exit Sub ────────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    If x < 0 Then Exit Sub
    x = 1
End Sub
`);
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'ExitSub: exit に到達可能');
    // exit への先行が 2 つ以上（Exit Sub パスと通常パス）
    assert.strictEqual(cfg.exit.preds.length >= 2, true, 'ExitSub: exit に複数パス');
    console.log('[PASS] Exit Sub');
}

// ─── GoTo（後向き参照） ───────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    x = 1
    GoTo Done
    x = 2
Done:
    x = 3
End Sub
`);
    // GoTo 後の x=2 を含むブロックから exit への経路がない（unreachable code）
    // ただし exit 自体は Done: ブロック経由で到達可能
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'GoTo: exit に到達可能');
    console.log('[PASS] GoTo（後向き参照）');
}

// ─── GoTo（前向き参照） ───────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
Forward:
    x = 1
    GoTo Forward
End Sub
`);
    // バックエッジが作られること（GoTo で後ろのラベルへジャンプ）
    let hasBackEdge = false;
    for (const b of cfg.blocks) {
        for (const s of b.succs) {
            if (s.id <= b.id) { hasBackEdge = true; break; }
        }
    }
    assert.strictEqual(hasBackEdge, true, 'GoTo(前向き): バックエッジ存在');
    console.log('[PASS] GoTo（前向き参照 = 後向きジャンプ）');
}

// ─── Select Case ─────────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Select Case x
        Case 1
            y = 10
        Case 2
            y = 20
        Case Else
            y = 0
    End Select
    z = y
End Sub
`);
    // select ブロック → 3 ケース + else + merge
    assert.strictEqual(cfg.blocks.length >= 5, true, 'Select: ブロック数 >= 5');
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'Select: exit に到達可能');
    console.log('[PASS] Select Case');
}

// ─── ネスト構造（For の中の If） ─────────────────────────────────────────────
{
    const cfg = buildFromSource(`
Sub Test()
    Dim i As Long
    For i = 1 To 10
        If i Mod 2 = 0 Then
            s = s + i
        End If
    Next i
End Sub
`);
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, 'Nested: exit に到達可能');
    // バックエッジが存在
    let hasBackEdge = false;
    for (const b of cfg.blocks) {
        for (const s of b.succs) {
            if (s.id < b.id) { hasBackEdge = true; break; }
        }
    }
    assert.strictEqual(hasBackEdge, true, 'Nested: バックエッジ存在');
    console.log('[PASS] ネスト（For + If）');
}

// ─── 空のプロシージャ ────────────────────────────────────────────────────────
{
    const cfg = buildFromSource(`Sub NoOp()\nEnd Sub`);
    // entry → first(empty) → exit
    assert.strictEqual(cfg.blocks.length, 3, '空: ブロック数 3');
    const ids = reachable(cfg);
    assert.strictEqual(ids.has(cfg.exit.id), true, '空: exit に到達可能');
    console.log('[PASS] 空のプロシージャ');
}

console.log('\n✅ CFG 構築: 全テスト通過');
