/**
 * Def-Use チェーン解析 — vba-analyzer 向け
 *
 * プロシージャ内の指定行範囲に対して、変数の定義（Def）と使用（Use）を追跡し
 * Extract Function リファクタリングに必要な引数情報を算出する。
 *
 * 出力:
 *   inputs  — 範囲外で定義され範囲内で使用される変数（ByVal 引数候補）
 *   outputs — 範囲内で定義され範囲外（後）で使用される変数（ByRef 引数候補）
 *   locals  — 範囲内だけで完結する変数（ローカル変数候補）
 *
 * 制限:
 *   分岐マージなし（線形スキャン）。CFG ベースの精密解析は Phase 4C 以降。
 */

import {
    ProcedureDeclaration,
    Statement,
    Expression,
    AssignmentStatement,
    SetStatement,
    ForStatement,
    ForEachStatement,
    Identifier,
} from './parser';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface DefUseResult {
    /** 範囲外で定義され、範囲内で使用される変数（関数の ByVal 引数候補） */
    inputs: string[];
    /** 範囲内で定義され、範囲外（後）で使用される変数（ByRef 引数候補） */
    outputs: string[];
    /** 範囲内だけで完結する変数（抽出後のローカル変数候補） */
    locals: string[];
}

interface VarEvent {
    varName: string;
    line: number;  // 1-based（パーサーの loc に合わせる）
    kind: 'def' | 'use';
}

// ─── 公開 API ─────────────────────────────────────────────────────────────────

/**
 * プロシージャ内の行範囲 [startLine, endLine]（1-based, inclusive）を解析し、
 * Extract Function に必要な変数分類を返す。
 */
export function analyzeDefUse(
    proc: ProcedureDeclaration,
    startLine: number,
    endLine: number,
): DefUseResult {
    const events = collectEvents(proc);

    // 変数名ごとにイベントをグループ化
    const byVar = new Map<string, VarEvent[]>();
    for (const ev of events) {
        const list = byVar.get(ev.varName) ?? [];
        list.push(ev);
        byVar.set(ev.varName, list);
    }

    const inputs: string[] = [];
    const outputs: string[] = [];
    const locals: string[] = [];

    for (const [varName, evs] of byVar) {
        const hasDefBefore  = evs.some(e => e.kind === 'def' && e.line < startLine);
        const hasUseInRange = evs.some(e => e.kind === 'use' && e.line >= startLine && e.line <= endLine);
        const hasDefInRange = evs.some(e => e.kind === 'def' && e.line >= startLine && e.line <= endLine);
        const hasUseAfter   = evs.some(e => e.kind === 'use' && e.line > endLine);
        const onlyInRange   = evs.every(e => e.line >= startLine && e.line <= endLine);

        // 範囲内の最初のイベントが USE なら外部値に依存 → input
        const firstInRange = evs.find(e => e.line >= startLine && e.line <= endLine);
        const firstIsUse   = firstInRange?.kind === 'use';

        const isInput  = hasUseInRange && (hasDefBefore || firstIsUse);
        const isOutput = hasDefInRange && hasUseAfter;

        if (isOutput) outputs.push(varName);
        if (isInput)  inputs.push(varName);
        if (!isInput && !isOutput && onlyInRange) locals.push(varName);
    }

    return {
        inputs:  inputs.sort(),
        outputs: outputs.sort(),
        locals:  locals.sort(),
    };
}

// ─── イベント収集 ─────────────────────────────────────────────────────────────

function collectEvents(proc: ProcedureDeclaration): VarEvent[] {
    const events: VarEvent[] = [];
    const procStartLine = proc.loc?.start.line ?? 1;

    // パラメーターはプロシージャ開始行で Def として登録
    for (const param of proc.parameters) {
        events.push({ varName: param.name.toLowerCase(), line: procStartLine, kind: 'def' });
    }

    for (const stmt of proc.body) {
        collectStmtEvents(stmt, events);
    }

    return events;
}

function collectStmtEvents(stmt: Statement, out: VarEvent[]): void {
    const line = (stmt as any).loc?.start.line ?? 0;

    switch (stmt.type) {
        case 'AssignmentStatement': {
            const as_ = stmt as AssignmentStatement;
            // RHS を先に収集（右辺が先に評価される）→ total = total + 1 で total が USE-before-DEF
            collectUseFromExpr(as_.right, line, out);
            collectDefFromLHS(as_.left, line, out);
            break;
        }
        case 'SetStatement': {
            const ss = stmt as SetStatement;
            collectUseFromExpr(ss.right, line, out);
            collectDefFromLHS(ss.left, line, out);
            break;
        }
        case 'ForStatement': {
            const fs = stmt as ForStatement;
            // ループカウンター変数は Def
            out.push({ varName: fs.identifier.name.toLowerCase(), line, kind: 'def' });
            collectUseFromExpr(fs.start, line, out);
            collectUseFromExpr(fs.end, line, out);
            if (fs.step) collectUseFromExpr(fs.step, line, out);
            for (const s of fs.body) collectStmtEvents(s, out);
            break;
        }
        case 'ForEachStatement': {
            const fe = stmt as ForEachStatement;
            out.push({ varName: fe.variable.name.toLowerCase(), line, kind: 'def' });
            collectUseFromExpr(fe.collection, line, out);
            for (const s of fe.body) collectStmtEvents(s, out);
            break;
        }
        case 'ReDimStatement': {
            const rd = stmt as any;
            out.push({ varName: rd.name.name.toLowerCase(), line, kind: 'def' });
            for (const b of (rd.bounds ?? [])) {
                if (b.lower) collectUseFromExpr(b.lower, line, out);
                if (b.upper) collectUseFromExpr(b.upper, line, out);
            }
            break;
        }
        case 'IfStatement': {
            const is = stmt as any;
            collectUseFromExpr(is.condition, line, out);
            const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
            const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
            for (const s of cons) collectStmtEvents(s, out);
            for (const s of alt)  collectStmtEvents(s, out);
            break;
        }
        case 'DoWhileStatement':
        case 'WhileStatement': {
            const ws = stmt as any;
            if (ws.condition) collectUseFromExpr(ws.condition, line, out);
            for (const s of (ws.body ?? [])) collectStmtEvents(s, out);
            break;
        }
        case 'WithStatement': {
            const ws = stmt as any;
            collectUseFromExpr(ws.object, line, out);
            for (const s of (ws.body ?? [])) collectStmtEvents(s, out);
            break;
        }
        case 'SelectCaseStatement': {
            const sc = stmt as any;
            collectUseFromExpr(sc.expression, line, out);
            for (const clause of (sc.cases ?? [])) {
                for (const s of (clause.body ?? [])) collectStmtEvents(s, out);
            }
            if (sc.elseBody) {
                for (const s of sc.elseBody) collectStmtEvents(s, out);
            }
            break;
        }
        case 'CallStatement': {
            collectUseFromExpr((stmt as any).expression, line, out);
            break;
        }
    }
}

function collectDefFromLHS(expr: Expression, line: number, out: VarEvent[]): void {
    if (!expr) return;
    if (expr.type === 'Identifier') {
        out.push({ varName: (expr as Identifier).name.toLowerCase(), line, kind: 'def' });
    }
    // MemberExpression の左辺（obj.Prop = ...）は obj が use
    if (expr.type === 'MemberExpression') {
        collectUseFromExpr((expr as any).object, line, out);
    }
}

function collectUseFromExpr(expr: Expression, line: number, out: VarEvent[]): void {
    if (!expr) return;

    switch (expr.type) {
        case 'Identifier':
            out.push({ varName: (expr as Identifier).name.toLowerCase(), line, kind: 'use' });
            break;
        case 'MemberExpression': {
            const me = expr as any;
            collectUseFromExpr(me.object, line, out);
            // me.property は識別子だが変数ではなくプロパティ名なので収集しない
            break;
        }
        case 'CallExpression': {
            const ce = expr as any;
            collectUseFromExpr(ce.callee, line, out);
            for (const arg of (ce.args ?? [])) collectUseFromExpr(arg, line, out);
            break;
        }
        case 'BinaryExpression':
        case 'LogicalExpression': {
            const be = expr as any;
            collectUseFromExpr(be.left, line, out);
            collectUseFromExpr(be.right, line, out);
            break;
        }
        case 'UnaryExpression':
            collectUseFromExpr((expr as any).argument, line, out);
            break;
        case 'ParenthesizedExpression':
            collectUseFromExpr((expr as any).expression, line, out);
            break;
    }
}
