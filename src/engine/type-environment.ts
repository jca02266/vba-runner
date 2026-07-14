/**
 * 静的型環境 (MS-VBAL §5.2 宣言ベース)
 *
 * Dim / Const / パラメーター / Function 戻り値 / Enum メンバーを
 * AST から収集し、識別子 → TypeInfo のマップを提供する。
 *
 * 対象: Phase 1（宣言ベース）。代入フロー追跡は FlowEnvironment (Phase 2) で行う。
 */

import {
    Program,
    Statement,
    Expression,
    VariableDeclaration,
    ConstDeclaration,
    ProcedureDeclaration,
    EnumDeclaration,
    ClassDeclaration,
    NumberLiteral,
    StringLiteral,
    UnaryExpression,
} from './parser';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export type TypeInfoKind = 'variable' | 'const' | 'parameter' | 'function' | 'enum-member';

export interface TypeInfo {
    kind: TypeInfoKind;
    /** Dim x As T の T。未指定なら 'Variant' */
    declaredType: string;
    /** Const の場合の即値（NumberLiteral / StringLiteral から抽出） */
    constValue?: number | string | boolean;
    /** 配列かどうか */
    isArray: boolean;
}

/** プロシージャスコープ内の型マップ */
export interface ProcedureTypeScope {
    /** 小文字化したプロシージャ名 */
    procName: string;
    /** 小文字化した識別子名 → TypeInfo */
    vars: Map<string, TypeInfo>;
}

/** モジュール全体の型環境 */
export interface ModuleTypeEnvironment {
    /** モジュールレベルの宣言 */
    moduleVars: Map<string, TypeInfo>;
    /** プロシージャごとのローカル型マップ */
    procedures: Map<string, ProcedureTypeScope>;
}

// ─── ユーティリティ ───────────────────────────────────────────────────────────

const NUMERIC_TYPES = new Set([
    'byte', 'integer', 'long', 'longlng', 'longlong', 'longptr', 'single',
    'double', 'currency', 'decimal',
]);

/** 型名が数値型かどうか */
export function isNumericType(declaredType?: string): boolean {
    if (!declaredType) return false;
    return NUMERIC_TYPES.has(declaredType.toLowerCase());
}

/** Const の右辺リテラルから即値を抽出する */
function extractConstValue(expr: Expression): number | string | boolean | undefined {
    if (expr.type === 'NumberLiteral') return (expr as NumberLiteral).value;
    if (expr.type === 'StringLiteral') return (expr as StringLiteral).value;
    // -1, -100 のような単項マイナス付き数値
    if (expr.type === 'UnaryExpression') {
        const ue = expr as UnaryExpression;
        if (ue.operator === '-' && ue.argument?.type === 'NumberLiteral') {
            return -((ue.argument as NumberLiteral).value);
        }
    }
    // True / False キーワードも定数として扱う
    if (expr.type === 'Identifier') {
        const name = (expr as any).name?.toLowerCase();
        if (name === 'true') return true;
        if (name === 'false') return false;
    }
    return undefined;
}

// ─── コア実装 ─────────────────────────────────────────────────────────────────

/**
 * モジュールのプログラム AST から型環境を構築する。
 *
 * 2パス構成:
 *   Pass 1 — モジュールレベルの宣言を収集
 *   Pass 2 — プロシージャ内の宣言・パラメーターを収集
 */
export function buildTypeEnvironment(program: Program): ModuleTypeEnvironment {
    const moduleVars = new Map<string, TypeInfo>();
    const procedures = new Map<string, ProcedureTypeScope>();

    // Pass 1: モジュールレベル
    for (const stmt of program.body) {
        collectModuleLevel(stmt, moduleVars);
    }

    // Pass 2: プロシージャ内
    for (const stmt of program.body) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            const scope = buildProcedureScope(proc);
            procedures.set(proc.name.name.toLowerCase(), scope);
        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            for (const proc of cls.procedures) {
                const scope = buildProcedureScope(proc);
                procedures.set(proc.name.name.toLowerCase(), scope);
            }
        }
    }

    return { moduleVars, procedures };
}

function collectModuleLevel(stmt: Statement, out: Map<string, TypeInfo>): void {
    switch (stmt.type) {
        case 'VariableDeclaration': {
            const vd = stmt as VariableDeclaration;
            for (const decl of vd.declarations) {
                const key = decl.name.name.toLowerCase();
                out.set(key, {
                    kind: 'variable',
                    declaredType: decl.objectType ?? 'Variant',
                    isArray: decl.isArray,
                });
            }
            break;
        }
        case 'ConstDeclaration': {
            const cd = stmt as ConstDeclaration;
            for (const decl of cd.declarations) {
                const key = decl.name.name.toLowerCase();
                out.set(key, {
                    kind: 'const',
                    declaredType: inferConstType(decl.value),
                    constValue: extractConstValue(decl.value),
                    isArray: false,
                });
            }
            break;
        }
        case 'ProcedureDeclaration': {
            const pd = stmt as ProcedureDeclaration;
            const key = pd.name.name.toLowerCase();
            if (pd.isFunction || pd.isProperty) {
                out.set(key, {
                    kind: 'function',
                    declaredType: pd.returnType ?? 'Variant',
                    isArray: false,
                });
            }
            break;
        }
        case 'EnumDeclaration': {
            const ed = stmt as EnumDeclaration;
            // Enum 名自体は型名として登録
            out.set(ed.name.name.toLowerCase(), {
                kind: 'variable',
                declaredType: 'Long',
                isArray: false,
            });
            // 各メンバーを Long 定数として登録
            let autoValue = 0;
            for (const member of ed.members) {
                const key = member.name.name.toLowerCase();
                let constValue: number | undefined;
                if (member.value) {
                    const extracted = extractConstValue(member.value);
                    if (typeof extracted === 'number') {
                        autoValue = extracted;
                        constValue = extracted;
                    }
                } else {
                    constValue = autoValue;
                }
                out.set(key, {
                    kind: 'enum-member',
                    declaredType: 'Long',
                    constValue,
                    isArray: false,
                });
                autoValue++;
            }
            break;
        }
        case 'ClassDeclaration': {
            const cls = stmt as ClassDeclaration;
            // クラスフィールドをモジュールレベルに登録
            for (const field of cls.fields) {
                collectModuleLevel(field, out);
            }
            break;
        }
    }
}

function buildProcedureScope(proc: ProcedureDeclaration): ProcedureTypeScope {
    const vars = new Map<string, TypeInfo>();

    // 戻り値変数（Function 名への代入 = 戻り値）
    if (proc.isFunction) {
        vars.set(proc.name.name.toLowerCase(), {
            kind: 'function',
            declaredType: proc.returnType ?? 'Variant',
            isArray: false,
        });
    }

    // パラメーター
    for (const param of proc.parameters) {
        vars.set(param.name.toLowerCase(), {
            kind: 'parameter',
            declaredType: param.paramType ?? 'Variant',
            isArray: param.isArray ?? false,
        });
    }

    // ボディ内のローカル宣言
    for (const stmt of proc.body) {
        collectLocalDeclarations(stmt, vars);
    }

    return { procName: proc.name.name.toLowerCase(), vars };
}

function collectLocalDeclarations(stmt: Statement, out: Map<string, TypeInfo>): void {
    switch (stmt.type) {
        case 'VariableDeclaration': {
            const vd = stmt as VariableDeclaration;
            for (const decl of vd.declarations) {
                out.set(decl.name.name.toLowerCase(), {
                    kind: 'variable',
                    declaredType: decl.objectType ?? 'Variant',
                    isArray: decl.isArray,
                });
            }
            break;
        }
        case 'ConstDeclaration': {
            const cd = stmt as ConstDeclaration;
            for (const decl of cd.declarations) {
                out.set(decl.name.name.toLowerCase(), {
                    kind: 'const',
                    declaredType: inferConstType(decl.value),
                    constValue: extractConstValue(decl.value),
                    isArray: false,
                });
            }
            break;
        }
        // ブロック文内の宣言も再帰的に収集
        case 'IfStatement': {
            const is = stmt as any;
            if (is.consequent) {
                const stmts = Array.isArray(is.consequent) ? is.consequent : [is.consequent];
                for (const s of stmts) collectLocalDeclarations(s, out);
            }
            if (is.alternate) {
                const stmts = Array.isArray(is.alternate) ? is.alternate : [is.alternate];
                for (const s of stmts) collectLocalDeclarations(s, out);
            }
            break;
        }
        case 'ForStatement':
        case 'ForEachStatement':
        case 'DoWhileStatement':
        case 'WhileStatement':
        case 'WithStatement': {
            const bs = stmt as any;
            const body: Statement[] = bs.body ?? [];
            for (const s of body) collectLocalDeclarations(s, out);
            break;
        }
        case 'SelectCaseStatement': {
            const sc = stmt as any;
            for (const clause of (sc.cases ?? [])) {
                for (const s of (clause.body ?? [])) collectLocalDeclarations(s, out);
            }
            break;
        }
    }
}

function inferConstType(expr: Expression): string {
    if (!expr) return 'Variant';
    if (expr.type === 'NumberLiteral') {
        const v = (expr as NumberLiteral).value;
        return Number.isInteger(v) ? 'Long' : 'Double';
    }
    if (expr.type === 'StringLiteral') return 'String';
    if (expr.type === 'UnaryExpression') {
        return inferConstType((expr as UnaryExpression).argument);
    }
    return 'Variant';
}

// ─── 照会 API ─────────────────────────────────────────────────────────────────

/**
 * 識別子の TypeInfo を返す。
 * procName を指定するとローカルスコープを優先して検索し、なければモジュールレベルを返す。
 */
export function lookupType(
    env: ModuleTypeEnvironment,
    name: string,
    procName?: string,
): TypeInfo | undefined {
    const key = name.toLowerCase();

    if (procName) {
        const scope = env.procedures.get(procName.toLowerCase());
        if (scope) {
            const local = scope.vars.get(key);
            if (local) return local;
        }
    }

    return env.moduleVars.get(key);
}
