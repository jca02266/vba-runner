/**
 * 制御フローグラフ（CFG）構築 — Phase 4A
 *
 * プロシージャの AST を基本ブロック列 + エッジのグラフに変換する。
 * 到達定義・生変数・デッドストア解析（Phase 4B〜D）の前提となる。
 *
 * 対応構文:
 *   順次実行 / If...Else / Select Case / For...Next / For Each...Next
 *   Do While...Loop / Do Until...Loop / While...Wend
 *   Exit For / Exit Do / Exit Sub / Exit Function
 *   GoTo（同一プロシージャ内ラベルジャンプ）
 */

import {
    ProcedureDeclaration,
    Statement,
} from './parser';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface BasicBlock {
    /** ブロック固有 ID（0 = entry, 1 = exit）*/
    id: number;
    /** このブロックに属するステートメント列 */
    stmts: Statement[];
    /** 後続ブロック */
    succs: BasicBlock[];
    /** 先行ブロック */
    preds: BasicBlock[];
    /** entry / exit / normal */
    kind: 'entry' | 'exit' | 'normal';
}

export interface CFG {
    /** 仮想エントリーブロック（空、本体先頭へのエッジを持つ） */
    entry: BasicBlock;
    /** 仮想出口ブロック（空、全終端からのエッジが集まる） */
    exit: BasicBlock;
    /** CFG 中の全ブロック（entry/exit を含む） */
    blocks: BasicBlock[];
}

// ─── 公開 API ─────────────────────────────────────────────────────────────────

export function buildCFG(proc: ProcedureDeclaration): CFG {
    const builder = new CFGBuilder();
    return builder.build(proc);
}

/**
 * CFG から到達不能な normal ブロック（stmts が 1 件以上）を返す。
 * entry からの BFS で到達可能集合を求め、残りを unreachable とみなす。
 */
export function findUnreachableBlocks(cfg: CFG): BasicBlock[] {
    const reachable = new Set<number>();
    const queue: BasicBlock[] = [cfg.entry];
    while (queue.length > 0) {
        const b = queue.pop()!;
        if (reachable.has(b.id)) continue;
        reachable.add(b.id);
        for (const s of b.succs) queue.push(s);
    }
    return cfg.blocks.filter(
        b => b.kind === 'normal' && !reachable.has(b.id) && b.stmts.length > 0,
    );
}

// ─── 内部実装 ─────────────────────────────────────────────────────────────────

class CFGBuilder {
    private idSeq = 0;
    readonly blocks: BasicBlock[] = [];

    private newBlock(kind: BasicBlock['kind'] = 'normal'): BasicBlock {
        const b: BasicBlock = { id: this.idSeq++, stmts: [], succs: [], preds: [], kind };
        this.blocks.push(b);
        return b;
    }

    private addEdge(from: BasicBlock, to: BasicBlock): void {
        if (!from.succs.includes(to)) from.succs.push(to);
        if (!to.preds.includes(from)) to.preds.push(from);
    }

    build(proc: ProcedureDeclaration): CFG {
        const entry = this.newBlock('entry');
        const exit  = this.newBlock('exit');

        // ラベルマップ（GoTo 解決用）
        const labelMap = new Map<string, BasicBlock>();
        // GoTo の前向き参照を後で解決するためのパッチリスト
        const gotoPatches: { from: BasicBlock; label: string }[] = [];
        // On Error GoTo のターゲットラベル（entry から到達可能とみなす）
        const onErrorLabels: string[] = [];

        const first = this.newBlock();
        this.addEdge(entry, first);

        const fallthroughs = this.buildStmts(
            proc.body, first, exit, null, null, labelMap, gotoPatches, onErrorLabels,
        );
        for (const b of fallthroughs) this.addEdge(b, exit);

        // GoTo 前向き参照の解決
        for (const patch of gotoPatches) {
            const target = labelMap.get(patch.label.toLowerCase());
            if (target) this.addEdge(patch.from, target);
        }

        // On Error GoTo のターゲットを entry から到達可能にする
        for (const label of onErrorLabels) {
            const target = labelMap.get(label.toLowerCase());
            if (target) this.addEdge(entry, target);
        }

        return { entry, exit, blocks: this.blocks };
    }

    /**
     * ステートメント列を処理し、フォールスルーするブロック群を返す。
     * @param current  現在追記中のブロック
     * @param procExit プロシージャ出口ブロック
     * @param loopNext ループ継続先（Continue 相当、VBA は Exit Do/For のみ）
     * @param loopExit ループ脱出先（Exit Do / Exit For の飛び先）
     */
    private buildStmts(
        stmts: Statement[],
        current: BasicBlock,
        procExit: BasicBlock,
        loopNext: BasicBlock | null,
        loopExit: BasicBlock | null,
        labelMap: Map<string, BasicBlock>,
        gotoPatches: { from: BasicBlock; label: string }[],
        onErrorLabels: string[] = [],
    ): BasicBlock[] {
        let cur = current;
        let terminated = false;

        for (const stmt of stmts) {
            const result = this.buildStmt(
                stmt, cur, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
            );
            if (result === null) {
                // 終端後も後続のラベル登録のため dead block を作って処理を続ける
                terminated = true;
                cur = this.newBlock();
            } else {
                terminated = false;
                cur = result;
            }
        }
        // 最後が終端なら外にフォールスルーしない
        return terminated ? [] : [cur];
    }

    /**
     * 単一ステートメントを処理する。
     * 戻り値: 次のステートメントを追記すべきブロック、または null（終端）。
     */
    private buildStmt(
        stmt: Statement,
        current: BasicBlock,
        procExit: BasicBlock,
        loopNext: BasicBlock | null,
        loopExit: BasicBlock | null,
        labelMap: Map<string, BasicBlock>,
        gotoPatches: { from: BasicBlock; label: string }[],
        onErrorLabels: string[] = [],
    ): BasicBlock | null {
        switch (stmt.type) {

            // ─── On Error GoTo ───────────────────────────────────────────
            case 'OnErrorStatement': {
                const errLabel = (stmt as any).label as string;
                if (errLabel && errLabel !== '0' && errLabel.toLowerCase() !== 'resume next') {
                    onErrorLabels.push(errLabel);
                }
                current.stmts.push(stmt);
                return current;
            }

            // ─── ラベル ──────────────────────────────────────────────────
            case 'LabelStatement': {
                const label = (stmt as any).label as string;
                // ラベル → 新しいブロックの先頭
                const labelBlock = this.newBlock();
                this.addEdge(current, labelBlock);
                labelMap.set(label.toLowerCase(), labelBlock);
                return labelBlock;
            }

            // ─── GoTo ────────────────────────────────────────────────────
            case 'GoToStatement': {
                const label = (stmt as any).label as string;
                const target = labelMap.get(label.toLowerCase());
                if (target) {
                    this.addEdge(current, target);
                } else {
                    // 前向き参照 → 後で解決
                    gotoPatches.push({ from: current, label });
                }
                return null; // GoTo 後はフォールスルーなし
            }

            // ─── Exit Sub / Exit Function ────────────────────────────────
            case 'ExitStatement': {
                const target = (stmt as any).target as string;
                if (target === 'Sub' || target === 'Function' || target === 'Property') {
                    current.stmts.push(stmt);
                    this.addEdge(current, procExit);
                    return null;
                }
                if ((target === 'For' || target === 'Do') && loopExit) {
                    current.stmts.push(stmt);
                    this.addEdge(current, loopExit);
                    return null;
                }
                // 未知の Exit → 終端扱い
                current.stmts.push(stmt);
                this.addEdge(current, procExit);
                return null;
            }

            // ─── If ──────────────────────────────────────────────────────
            case 'IfStatement': {
                const is = stmt as any;
                current.stmts.push(stmt); // 条件式をこのブロックに含める

                const thenBlock  = this.newBlock();
                const mergeBlock = this.newBlock();

                this.addEdge(current, thenBlock);

                const thenBody = Array.isArray(is.consequent)
                    ? is.consequent : (is.consequent ? [is.consequent] : []);
                const thenFalls = this.buildStmts(
                    thenBody, thenBlock, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
                );
                for (const b of thenFalls) this.addEdge(b, mergeBlock);

                const altBody = Array.isArray(is.alternate)
                    ? is.alternate : (is.alternate ? [is.alternate] : []);
                if (altBody.length > 0) {
                    const elseBlock = this.newBlock();
                    this.addEdge(current, elseBlock);
                    const elseFalls = this.buildStmts(
                        altBody, elseBlock, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
                    );
                    for (const b of elseFalls) this.addEdge(b, mergeBlock);
                } else {
                    // Else なし → 条件が偽のときは直接 merge
                    this.addEdge(current, mergeBlock);
                }

                return mergeBlock;
            }

            // ─── Select Case ─────────────────────────────────────────────
            case 'SelectCaseStatement': {
                const sc = stmt as any;
                current.stmts.push(stmt);
                const mergeBlock = this.newBlock();

                for (const clause of (sc.cases ?? [])) {
                    const caseBlock = this.newBlock();
                    this.addEdge(current, caseBlock);
                    const falls = this.buildStmts(
                        clause.body ?? [], caseBlock, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
                    );
                    for (const b of falls) this.addEdge(b, mergeBlock);
                }

                if (sc.elseBody && sc.elseBody.length > 0) {
                    const elseBlock = this.newBlock();
                    this.addEdge(current, elseBlock);
                    const falls = this.buildStmts(
                        sc.elseBody, elseBlock, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
                    );
                    for (const b of falls) this.addEdge(b, mergeBlock);
                } else {
                    this.addEdge(current, mergeBlock);
                }

                return mergeBlock;
            }

            // ─── For...Next ──────────────────────────────────────────────
            case 'ForStatement': {
                const fs = stmt as any;
                // ヘッダーブロック（初期化 + 条件判定）
                current.stmts.push(stmt);
                const bodyBlock  = this.newBlock();
                const afterBlock = this.newBlock();

                this.addEdge(current, bodyBlock);   // ループ入る
                this.addEdge(current, afterBlock);  // 初回から条件偽（空配列など）

                const bodyFalls = this.buildStmts(
                    fs.body ?? [], bodyBlock, procExit, current, afterBlock, labelMap, gotoPatches, onErrorLabels,
                );
                for (const b of bodyFalls) this.addEdge(b, current); // バックエッジ

                return afterBlock;
            }

            // ─── For Each...Next ─────────────────────────────────────────
            case 'ForEachStatement': {
                const fe = stmt as any;
                current.stmts.push(stmt);
                const bodyBlock  = this.newBlock();
                const afterBlock = this.newBlock();

                this.addEdge(current, bodyBlock);
                this.addEdge(current, afterBlock);

                const bodyFalls = this.buildStmts(
                    fe.body ?? [], bodyBlock, procExit, current, afterBlock, labelMap, gotoPatches, onErrorLabels,
                );
                for (const b of bodyFalls) this.addEdge(b, current);

                return afterBlock;
            }

            // ─── Do While / Do Until ─────────────────────────────────────
            case 'DoWhileStatement': {
                const dw = stmt as any;
                const headerBlock = this.newBlock();
                const bodyBlock   = this.newBlock();
                const afterBlock  = this.newBlock();

                this.addEdge(current, headerBlock);
                headerBlock.stmts.push(stmt); // 条件式
                this.addEdge(headerBlock, bodyBlock);
                this.addEdge(headerBlock, afterBlock); // 条件偽

                const bodyFalls = this.buildStmts(
                    dw.body ?? [], bodyBlock, procExit, headerBlock, afterBlock, labelMap, gotoPatches, onErrorLabels,
                );
                for (const b of bodyFalls) this.addEdge(b, headerBlock);

                return afterBlock;
            }

            // ─── While...Wend ────────────────────────────────────────────
            case 'WhileStatement': {
                const ws = stmt as any;
                const headerBlock = this.newBlock();
                const bodyBlock   = this.newBlock();
                const afterBlock  = this.newBlock();

                this.addEdge(current, headerBlock);
                headerBlock.stmts.push(stmt);
                this.addEdge(headerBlock, bodyBlock);
                this.addEdge(headerBlock, afterBlock);

                const bodyFalls = this.buildStmts(
                    ws.body ?? [], bodyBlock, procExit, headerBlock, afterBlock, labelMap, gotoPatches, onErrorLabels,
                );
                for (const b of bodyFalls) this.addEdge(b, headerBlock);

                return afterBlock;
            }

            // ─── With ────────────────────────────────────────────────────
            case 'WithStatement': {
                const ws = stmt as any;
                current.stmts.push(stmt);
                const bodyFalls = this.buildStmts(
                    ws.body ?? [], current, procExit, loopNext, loopExit, labelMap, gotoPatches, onErrorLabels,
                );
                if (bodyFalls.length === 0) return null;
                // With は分岐しないので bodyFalls は [current] または []
                return bodyFalls[bodyFalls.length - 1];
            }

            // ─── その他（単純ステートメント） ────────────────────────────
            default:
                current.stmts.push(stmt);
                return current;
        }
    }
}
