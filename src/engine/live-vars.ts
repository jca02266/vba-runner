/**
 * 生変数解析（Live Variable Analysis）— Phase 4C
 *
 * CFG の各基本ブロックについて「その点で生きている（将来使われる可能性がある）変数」の
 * 集合を算出する後向きデータフロー解析。
 *
 * IN[B]  = use[B] ∪ (OUT[B] − def[B])
 * OUT[B] = ⋃ IN[S]  （S: B の後継ブロック）
 *
 * デッドストア検出（Phase 4D）の前段として使用する。
 */

import { CFG } from './cfg';
import {
    Statement,
    AssignmentStatement,
    SetStatement,
    ForStatement,
    ForEachStatement,
} from './parser';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface LiveVarsResult {
    /** ブロック ID → そのブロックの先頭で生きている変数集合 */
    blockIn:  Map<number, Set<string>>;
    /** ブロック ID → そのブロックの末尾で生きている変数集合 */
    blockOut: Map<number, Set<string>>;
}

// ─── 公開 API ─────────────────────────────────────────────────────────────────

/**
 * CFG 全体に対して生変数解析を実行する。
 * @param extraLiveAtExit 出口ブロックで常に生きているとみなす変数名集合
 *                        （関数戻り値変数や ByRef 引数など）
 */
export function computeLiveVars(
    cfg: CFG,
    extraLiveAtExit: ReadonlySet<string> = new Set(),
): LiveVarsResult {
    const blockIn  = new Map<number, Set<string>>();
    const blockOut = new Map<number, Set<string>>();
    for (const b of cfg.blocks) {
        blockIn.set(b.id,  new Set());
        blockOut.set(b.id, new Set());
    }

    // use[B]: B 内でその前に def されずに参照される変数
    // def[B]: B 内で最初に def される変数
    const useMap = new Map<number, Set<string>>();
    const defMap = new Map<number, Set<string>>();
    for (const b of cfg.blocks) {
        const use = new Set<string>();
        const def = new Set<string>();
        for (const stmt of b.stmts) {
            for (const v of getStmtUses(stmt)) {
                if (!def.has(v)) use.add(v);
            }
            for (const v of getStmtDefs(stmt)) {
                def.add(v);
            }
        }
        useMap.set(b.id, use);
        defMap.set(b.id, def);
    }

    // 後向きワークリスト
    const worklist = new Set<number>(cfg.blocks.map(b => b.id));

    while (worklist.size > 0) {
        const bid = worklist.values().next().value as number;
        worklist.delete(bid);
        const b = cfg.blocks.find(bl => bl.id === bid)!;

        // OUT[B] = ⋃ IN[S]
        const newOut = new Set<string>();
        for (const succ of b.succs) {
            for (const v of blockIn.get(succ.id)!) newOut.add(v);
        }
        // 出口ブロックは常に生きている変数を追加
        if (b.id === cfg.exit.id) {
            for (const v of extraLiveAtExit) newOut.add(v);
        }
        blockOut.set(b.id, newOut);

        // IN[B] = use[B] ∪ (OUT[B] − def[B])
        const newIn = new Set<string>(useMap.get(b.id)!);
        const bDef  = defMap.get(b.id)!;
        for (const v of newOut) {
            if (!bDef.has(v)) newIn.add(v);
        }

        const oldIn = blockIn.get(b.id)!;
        if (!setsEqual(oldIn, newIn)) {
            blockIn.set(b.id, newIn);
            // IN[B] が変わった → B の先行ブロックの OUT に影響
            for (const pred of b.preds) worklist.add(pred.id);
        }
    }

    return { blockIn, blockOut };
}

// ─── ステートメントから def/use を抽出 ──────────────────────────────────────

/**
 * ステートメントが定義（書き込み）する変数名を返す。
 * CFG ブロックのヘッダーレベルのみ（子ブロックは別途解析）。
 */
export function getStmtDefs(stmt: Statement): string[] {
    switch (stmt.type) {
        case 'AssignmentStatement': {
            const as = stmt as AssignmentStatement;
            if (as.left.type === 'Identifier')
                return [(as.left as any).name.toLowerCase()];
            return [];
        }
        case 'SetStatement': {
            const ss = stmt as SetStatement;
            if (ss.left.type === 'Identifier')
                return [(ss.left as any).name.toLowerCase()];
            return [];
        }
        case 'ForStatement':
            return [(stmt as ForStatement).identifier.name.toLowerCase()];
        case 'ForEachStatement':
            return [(stmt as ForEachStatement).variable.name.toLowerCase()];
        case 'ReDimStatement':
            return [(stmt as any).name.name.toLowerCase()];
        default:
            return [];
    }
}

/**
 * ステートメントが参照（読み取り）する変数名を返す。
 * CFG ブロックのヘッダーレベルのみ（子ブロックは別途解析）。
 */
export function getStmtUses(stmt: Statement): string[] {
    const uses = new Set<string>();
    switch (stmt.type) {
        case 'AssignmentStatement': {
            const as = stmt as AssignmentStatement;
            collectLhsUses((as as any).left, uses);
            collectExprUses((as as any).right, uses);
            break;
        }
        case 'SetStatement': {
            const ss = stmt as SetStatement;
            collectLhsUses((ss as any).left, uses);
            collectExprUses((ss as any).right, uses);
            break;
        }
        case 'ForStatement': {
            const fs = stmt as ForStatement;
            collectExprUses(fs.start as any, uses);
            collectExprUses(fs.end as any, uses);
            if (fs.step) collectExprUses(fs.step as any, uses);
            break;
        }
        case 'ForEachStatement':
            collectExprUses((stmt as ForEachStatement).collection as any, uses);
            break;
        case 'IfStatement':
            collectExprUses((stmt as any).condition, uses);
            break;
        case 'DoWhileStatement':
        case 'WhileStatement':
            if ((stmt as any).condition) collectExprUses((stmt as any).condition, uses);
            break;
        case 'SelectCaseStatement':
            collectExprUses((stmt as any).expression, uses);
            break;
        case 'CallStatement':
            collectExprUses((stmt as any).expression, uses);
            break;
        case 'WithStatement':
            collectExprUses((stmt as any).object, uses);
            break;
        case 'ReDimStatement':
            for (const b of ((stmt as any).bounds ?? [])) {
                if (b.lower) collectExprUses(b.lower, uses);
                if (b.upper) collectExprUses(b.upper, uses);
            }
            break;
    }
    return [...uses];
}

// ─── 内部ユーティリティ ───────────────────────────────────────────────────────

function collectLhsUses(expr: any, out: Set<string>): void {
    if (!expr) return;
    if (expr.type === 'Identifier') return; // 単純変数: DEF なので USE しない
    if (expr.type === 'MemberExpression') {
        collectExprUses(expr.object, out);
    }
    if (expr.type === 'CallExpression') {
        // arr(i) = x の形式: arr と i は USE
        if (expr.callee?.type === 'Identifier') out.add(expr.callee.name.toLowerCase());
        else collectExprUses(expr.callee, out);
        for (const arg of (expr.args ?? [])) collectExprUses(arg, out);
    }
}

function collectExprUses(expr: any, out: Set<string>): void {
    if (!expr) return;
    switch (expr.type) {
        case 'Identifier':
            out.add(expr.name.toLowerCase());
            break;
        case 'MemberExpression':
            collectExprUses(expr.object, out);
            break;
        case 'CallExpression':
            if (expr.callee?.type === 'Identifier') out.add(expr.callee.name.toLowerCase());
            else collectExprUses(expr.callee, out);
            for (const arg of (expr.args ?? [])) collectExprUses(arg, out);
            break;
        case 'BinaryExpression':
        case 'LogicalExpression':
            collectExprUses(expr.left, out);
            collectExprUses(expr.right, out);
            break;
        case 'UnaryExpression':
            collectExprUses(expr.argument, out);
            break;
        case 'ParenthesizedExpression':
            collectExprUses(expr.expression, out);
            break;
    }
}

function setsEqual(a: Set<string>, b: Set<string>): boolean {
    if (a.size !== b.size) return false;
    for (const v of a) if (!b.has(v)) return false;
    return true;
}
