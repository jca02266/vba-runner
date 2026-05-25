/**
 * 代入フロー環境 — Phase 2
 *
 * プロシージャ本体の SetStatement / AssignmentStatement を前向き走査し、
 * 変数に流れ込む型を推定する。TypeEnvironment (Phase 1) の宣言型を基点として、
 * 代入によって型を絞り込む。
 *
 * 制限:
 * - 分岐・ループでの型マージは行わない（保守的: 最後の代入の型を使用）
 * - 関数呼び出しの戻り型は KNOWN_METHOD_RETURN_TYPES で既知のもののみ解決
 * - 手続き間解析は Phase 4
 */

import {
    ProcedureDeclaration,
    Statement,
    Expression,
    SetStatement,
    AssignmentStatement,
    NumberLiteral,
    StringLiteral,
    NewExpression,
    MemberExpression,
    CallExpression,
    Identifier,
    ForStatement,
    ForEachStatement,
    DoWhileStatement,
    WhileStatement,
    WithStatement,
    IfStatement,
    SelectCaseStatement,
} from './parser';
import {
    ModuleTypeEnvironment,
    lookupType,
    isNumericType,
} from './type-environment';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface FlowTypeInfo {
    /** 推定された型名 */
    inferredType: string;
    /** 型の出どころ */
    source: 'declaration' | 'assignment' | 'propagation';
}

/** プロシージャスコープ内の変数名 → 推定型マップ */
export type FlowEnvironment = Map<string, FlowTypeInfo>;

// ─── 既知メソッド/プロパティの戻り型 ─────────────────────────────────────────

/**
 * Excel オブジェクトの既知メソッド・プロパティ → 戻り型マップ。
 * `obj.Property` / `obj.Method(...)` の右辺推定に使用。
 * キーは小文字。
 */
const KNOWN_PROPERTY_RETURN_TYPES: ReadonlyMap<string, string> = new Map([
    // Range 系
    ['range',       'Range'],
    ['cells',       'Range'],
    ['rows',        'Range'],
    ['columns',     'Range'],
    ['offset',      'Range'],
    ['resize',      'Range'],
    ['entirerow',   'Range'],
    ['entirecolumn','Range'],
    ['currentregion','Range'],
    ['usedrange',   'Range'],
    ['end',         'Range'],
    ['item',        'Range'],
    // Worksheet 系
    ['sheets',      'Worksheet'],
    ['worksheets',  'Worksheet'],
    ['activesheet', 'Worksheet'],
    // Workbook 系
    ['workbooks',   'Workbook'],
    ['activeworkbook','Workbook'],
    ['thisworkbook','Workbook'],
    // Application
    ['application', 'Application'],
]);

// ─── RHS 型推定 ───────────────────────────────────────────────────────────────

/**
 * 代入右辺式から型を推定する。
 * flowEnv には現時点のフロー環境を渡す（コピー伝播に使用）。
 */
export function inferExprType(
    expr: Expression,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: FlowEnvironment,
    procName?: string,
): string {
    if (!expr) return 'Variant';

    switch (expr.type) {
        case 'NumberLiteral': {
            const v = (expr as NumberLiteral).value;
            return Number.isInteger(v) ? 'Long' : 'Double';
        }
        case 'StringLiteral':
            return 'String';

        case 'NewExpression':
            return (expr as NewExpression).className;

        case 'Identifier': {
            const name = (expr as Identifier).name.toLowerCase();
            // フロー環境を優先（代入で絞り込まれた型）
            const flow = flowEnv.get(name);
            if (flow) return flow.inferredType;
            // 宣言型にフォールバック
            const info = lookupType(typeEnv, name, procName);
            if (info) return info.declaredType;
            return 'Variant';
        }

        case 'MemberExpression': {
            // obj.Property → KNOWN_PROPERTY_RETURN_TYPES でプロパティ名を引く
            const me = expr as MemberExpression;
            const prop = me.property.name.toLowerCase();
            const known = KNOWN_PROPERTY_RETURN_TYPES.get(prop);
            if (known) return known;
            return 'Variant';
        }

        case 'CallExpression': {
            const ce = expr as CallExpression;
            if (ce.callee.type === 'MemberExpression') {
                // obj.Method(...) → メソッド名で検索
                const me = ce.callee as MemberExpression;
                const method = me.property.name.toLowerCase();
                const known = KNOWN_PROPERTY_RETURN_TYPES.get(method);
                if (known) return known;
            } else if (ce.callee.type === 'Identifier') {
                // CreateObject("Excel.Range") 等
                const fnName = (ce.callee as Identifier).name.toLowerCase();
                if (fnName === 'createobject' && ce.args.length > 0) {
                    const arg0 = ce.args[0];
                    if (arg0.type === 'StringLiteral') {
                        return resolveProgIdType((arg0 as StringLiteral).value);
                    }
                }
            }
            // TypeEnv の関数戻り型
            if (ce.callee.type === 'Identifier') {
                const fnName = (ce.callee as Identifier).name;
                const info = lookupType(typeEnv, fnName, procName);
                if (info?.kind === 'function') return info.declaredType;
            }
            return 'Variant';
        }

        default:
            return 'Variant';
    }
}

/** ProgID 文字列から型名を推定する（CreateObject 向け） */
function resolveProgIdType(progId: string): string {
    const lower = progId.toLowerCase();
    if (lower.includes('range'))      return 'Range';
    if (lower.includes('worksheet'))  return 'Worksheet';
    if (lower.includes('workbook'))   return 'Workbook';
    if (lower.includes('dictionary')) return 'Dictionary';
    if (lower.includes('filesystem')) return 'FileSystemObject';
    return 'Object';
}

// ─── フロー環境の構築 ─────────────────────────────────────────────────────────

/**
 * プロシージャ本体を前向き走査し、FlowEnvironment を構築する。
 *
 * 初期値として TypeEnvironment の宣言型をコピーし、
 * Set / Assignment で上書きする。
 */
export function buildFlowEnvironment(
    proc: ProcedureDeclaration,
    typeEnv: ModuleTypeEnvironment,
): FlowEnvironment {
    const flowEnv: FlowEnvironment = new Map();
    const procName = proc.name.name;

    // 宣言型を初期値としてフロー環境に投入
    const procScope = typeEnv.procedures.get(procName.toLowerCase());
    if (procScope) {
        for (const [key, info] of procScope.vars) {
            if (info.declaredType !== 'Variant') {
                flowEnv.set(key, { inferredType: info.declaredType, source: 'declaration' });
            }
        }
    }

    // ボディを前向き走査
    for (const stmt of proc.body) {
        scanStatement(stmt, typeEnv, flowEnv, procName);
    }

    return flowEnv;
}

function scanStatement(
    stmt: Statement,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: FlowEnvironment,
    procName: string,
): void {
    switch (stmt.type) {
        case 'SetStatement': {
            const ss = stmt as SetStatement;
            updateFlowFromAssignment(ss.left, ss.right, typeEnv, flowEnv, procName);
            break;
        }
        case 'AssignmentStatement': {
            const as_ = stmt as AssignmentStatement;
            updateFlowFromAssignment(as_.left, as_.right, typeEnv, flowEnv, procName);
            break;
        }
        // ブロック文を再帰走査（分岐マージなし — 保守的に全パスをスキャン）
        case 'IfStatement': {
            const is = stmt as IfStatement;
            const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
            const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
            for (const s of cons) scanStatement(s, typeEnv, flowEnv, procName);
            for (const s of alt)  scanStatement(s, typeEnv, flowEnv, procName);
            break;
        }
        case 'ForStatement':
        case 'ForEachStatement':
        case 'DoWhileStatement':
        case 'WhileStatement':
        case 'WithStatement': {
            const bs = stmt as any;
            for (const s of (bs.body ?? [])) scanStatement(s, typeEnv, flowEnv, procName);
            break;
        }
        case 'SelectCaseStatement': {
            for (const clause of (stmt as any).cases ?? []) {
                for (const s of (clause.body ?? [])) scanStatement(s, typeEnv, flowEnv, procName);
            }
            break;
        }
    }
}

function updateFlowFromAssignment(
    left: Expression,
    right: Expression,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: FlowEnvironment,
    procName: string,
): void {
    if (!left || left.type !== 'Identifier') return;
    const varName = (left as Identifier).name.toLowerCase();
    const inferredType = inferExprType(right, typeEnv, flowEnv, procName);
    if (inferredType !== 'Variant') {
        flowEnv.set(varName, { inferredType, source: 'assignment' });
    }
}

// ─── 型照会ヘルパー ───────────────────────────────────────────────────────────

/**
 * TypeEnv + FlowEnv を統合して変数の型を返す。
 * FlowEnv（代入で絞り込まれた型）が宣言型より優先。
 */
export function resolveVarType(
    name: string,
    typeEnv: ModuleTypeEnvironment,
    flowEnv: FlowEnvironment,
    procName?: string,
): string {
    const lower = name.toLowerCase();
    const flow = flowEnv.get(lower);
    if (flow) return flow.inferredType;
    const info = lookupType(typeEnv, lower, procName);
    return info?.declaredType ?? 'Variant';
}
