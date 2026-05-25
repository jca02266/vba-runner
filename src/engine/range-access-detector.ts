/**
 * Range 変数経由のアクセス検出 — Phase 3
 *
 * TypeEnvironment + FlowEnvironment を使い、Range 型変数経由の
 * メソッド呼び出し・添字アクセスをソースコード内で特定する。
 *
 * 検出対象:
 *   rng.Item(row, col)   — MemberExpression + CallExpression
 *   rng(row, col)        — CallExpression (callee = Identifier)
 *   rng.Value            — MemberExpression（代入先・参照元）
 *
 * 使用例:
 *   const env  = buildTypeEnvironment(ast);
 *   const hits = detectRangeAccess(ast, env);
 */

import {
    Program,
    Statement,
    Expression,
    ProcedureDeclaration,
    CallExpression,
    MemberExpression,
    AssignmentStatement,
    SetStatement,
    Identifier,
    ClassDeclaration,
} from './parser';
import { buildTypeEnvironment, ModuleTypeEnvironment } from './type-environment';
import { buildFlowEnvironment, resolveVarType } from './flow-environment';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface RangeAccessHit {
    /** アクセスしている変数名 */
    varName: string;
    /** アクセス種別 */
    kind: 'member-call' | 'index-call' | 'member-access';
    /** アクセスしているプロパティ/メソッド名（index-call の場合は undefined） */
    property?: string;
    /** ソース上の行番号 (0-based) */
    line: number;
    /** ソース上の列番号 (0-based) */
    column: number;
}

// ─── コア実装 ─────────────────────────────────────────────────────────────────

/**
 * プログラム全体から Range 変数経由のアクセスを検出する。
 */
export function detectRangeAccess(
    program: Program,
    typeEnv?: ModuleTypeEnvironment,
): RangeAccessHit[] {
    const env = typeEnv ?? buildTypeEnvironment(program);
    const hits: RangeAccessHit[] = [];

    for (const stmt of program.body) {
        if (stmt.type === 'ProcedureDeclaration') {
            scanProcedure(stmt as ProcedureDeclaration, env, hits);
        } else if (stmt.type === 'ClassDeclaration') {
            for (const proc of (stmt as ClassDeclaration).procedures) {
                scanProcedure(proc, env, hits);
            }
        }
    }
    return hits;
}

function scanProcedure(
    proc: ProcedureDeclaration,
    typeEnv: ModuleTypeEnvironment,
    hits: RangeAccessHit[],
): void {
    const flowEnv = buildFlowEnvironment(proc, typeEnv);
    const procName = proc.name.name;

    for (const stmt of proc.body) {
        scanStatement(stmt, typeEnv, flowEnv, procName, hits);
    }
}

function scanStatement(
    stmt: Statement,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: import('./flow-environment').FlowEnvironment,
    procName: string,
    hits: RangeAccessHit[],
): void {
    switch (stmt.type) {
        case 'AssignmentStatement': {
            const as_ = stmt as AssignmentStatement;
            scanExpr(as_.left,  typeEnv, flowEnv, procName, hits);
            scanExpr(as_.right, typeEnv, flowEnv, procName, hits);
            break;
        }
        case 'SetStatement': {
            const ss = stmt as SetStatement;
            scanExpr(ss.right, typeEnv, flowEnv, procName, hits);
            break;
        }
        case 'IfStatement': {
            const is = stmt as any;
            const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
            const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
            scanExpr(is.condition, typeEnv, flowEnv, procName, hits);
            for (const s of cons) scanStatement(s, typeEnv, flowEnv, procName, hits);
            for (const s of alt)  scanStatement(s, typeEnv, flowEnv, procName, hits);
            break;
        }
        case 'ForStatement':
        case 'ForEachStatement':
        case 'DoWhileStatement':
        case 'WhileStatement':
        case 'WithStatement': {
            const bs = stmt as any;
            for (const s of (bs.body ?? [])) scanStatement(s, typeEnv, flowEnv, procName, hits);
            break;
        }
        case 'SelectCaseStatement': {
            for (const clause of (stmt as any).cases ?? []) {
                for (const s of (clause.body ?? [])) scanStatement(s, typeEnv, flowEnv, procName, hits);
            }
            break;
        }
        case 'CallStatement': {
            scanExpr((stmt as any).expression, typeEnv, flowEnv, procName, hits);
            break;
        }
    }
}

function scanExpr(
    expr: Expression,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: import('./flow-environment').FlowEnvironment,
    procName: string,
    hits: RangeAccessHit[],
): void {
    if (!expr) return;

    switch (expr.type) {
        case 'CallExpression': {
            const ce = expr as CallExpression;

            if (ce.callee.type === 'MemberExpression') {
                // rng.Item(r, c) / rng.Value = ...
                const me = ce.callee as MemberExpression;
                if (me.object.type === 'Identifier') {
                    const varName = (me.object as Identifier).name;
                    const t = resolveVarType(varName, typeEnv, flowEnv, procName);
                    if (isRangeType(t)) {
                        hits.push({
                            varName,
                            kind: 'member-call',
                            property: me.property.name,
                            line: me.object.loc?.start.line ?? 0,
                            column: me.object.loc?.start.column ?? 0,
                        });
                    }
                }
            } else if (ce.callee.type === 'Identifier') {
                // rng(r, c) — 添字アクセス形式
                const varName = (ce.callee as Identifier).name;
                const t = resolveVarType(varName, typeEnv, flowEnv, procName);
                if (isRangeType(t) && ce.args.length > 0) {
                    hits.push({
                        varName,
                        kind: 'index-call',
                        line: ce.callee.loc?.start.line ?? 0,
                        column: ce.callee.loc?.start.column ?? 0,
                    });
                }
            }

            // 引数も再帰スキャン
            for (const arg of ce.args) scanExpr(arg, typeEnv, flowEnv, procName, hits);
            break;
        }

        case 'MemberExpression': {
            // rng.Value（呼び出しでない単独のメンバーアクセス）
            const me = expr as MemberExpression;
            if (me.object.type === 'Identifier') {
                const varName = (me.object as Identifier).name;
                const t = resolveVarType(varName, typeEnv, flowEnv, procName);
                if (isRangeType(t)) {
                    hits.push({
                        varName,
                        kind: 'member-access',
                        property: me.property.name,
                        line: me.object.loc?.start.line ?? 0,
                        column: me.object.loc?.start.column ?? 0,
                    });
                }
            }
            scanExpr(me.object, typeEnv, flowEnv, procName, hits);
            break;
        }

        case 'BinaryExpression':
        case 'LogicalExpression': {
            const be = expr as any;
            scanExpr(be.left,  typeEnv, flowEnv, procName, hits);
            scanExpr(be.right, typeEnv, flowEnv, procName, hits);
            break;
        }

        case 'UnaryExpression': {
            scanExpr((expr as any).argument, typeEnv, flowEnv, procName, hits);
            break;
        }

        case 'ParenthesizedExpression': {
            scanExpr((expr as any).expression, typeEnv, flowEnv, procName, hits);
            break;
        }
    }
}

function isRangeType(typeName: string): boolean {
    return typeName.toLowerCase() === 'range';
}
