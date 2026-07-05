/**
 * Variant 型変数の型推論
 *
 * カーソルのある手続きに限定して解析し、その手続きが呼ぶ関数は
 * メモ化付きで再帰的に辿る（最大 MAX_DEPTH 段）。
 *
 * 推論ルール:
 *   NumberLiteral（整数）→ Long
 *   NumberLiteral（小数）→ Double
 *   StringLiteral        → String
 *   DateLiteral          → Date
 *   True / False         → Boolean
 *   Nothing              → Object
 *   &演算子              → String
 *   算術演算子           → Long / Double（両辺から判定）
 *   関数呼び出し         → その関数の宣言型 or 再帰推論
 *   その他               → 推論不能（Variant のまま）
 */

import {
    Statement,
    Expression,
    ProcedureDeclaration,
    VariableDeclaration,
    Identifier,
    NumberLiteral,
    StringLiteral,
    CallExpression,
} from '../engine/parser';
import { Evaluator } from '../engine/evaluator';

const MAX_DEPTH = 3;

export type InferredType = string | null; // null = 推論不能

export interface VarHint {
    /** 変数名（小文字） */
    varName: string;
    /** 推論された型名 */
    inferredType: InferredType;
    /** ヒント表示位置 (1-based) */
    line: number;
    column: number;
    endColumn: number;
    /** ヒントの種別 */
    kind: 'var' | 'param' | 'return';
}

// ─── 公開エントリーポイント ───────────────────────────────────────────────────

/**
 * `proc` 内の Variant/型なし変数を推論し VarHint[] を返す。
 * `allProcs` は同ファイルの全手続きマップ（名前小文字→宣言）。
 * `memo` は呼び出し元で生成し、複数呼び出し間で共有することで効率化できる。
 */
export function inferVariantTypes(
    proc: ProcedureDeclaration,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType> = new Map(),
): VarHint[] {
    // 1. Variant / 型なし宣言を収集
    const variantVars = new Map<string, { line: number; column: number; endColumn: number }>();
    for (const stmt of proc.body) {
        if (stmt.type !== 'VariableDeclaration') continue;
        for (const decl of (stmt as VariableDeclaration).declarations) {
            const t = decl.objectType?.toLowerCase();
            if (t && t !== 'variant' && t !== 'object') continue; // 明示的な非 Variant/Object 型はスキップ
            const loc = (decl.name as any).loc;
            if (!loc) continue;
            const nameEnd = loc.end?.column ?? loc.start.column + decl.name.name.length;
            variantVars.set(decl.name.name.toLowerCase(), {
                line:      loc.start.line,
                column:    loc.start.column,
                endColumn: (decl as any).arrayEndColumn ?? nameEnd,
            });
        }
    }
    if (variantVars.size === 0) return [];

    // 2. 手続きボディを走査して最初の代入から型を推論
    const resolved = new Map<string, InferredType>();
    collectAssignmentTypes(proc.body, variantVars, resolved, allProcs, memo, 0);

    // 3. VarHint に変換（宣言型と同じ型は冗長なのでスキップ）
    const hints: VarHint[] = [];
    // 変数ごとの宣言型を逆引きするマップを構築
    const declaredTypeMap = new Map<string, string>();
    for (const stmt of proc.body) {
        if (stmt.type !== 'VariableDeclaration') continue;
        for (const decl of (stmt as VariableDeclaration).declarations) {
            const t = decl.objectType?.toLowerCase();
            if (t) declaredTypeMap.set(decl.name.name.toLowerCase(), t);
        }
    }
    for (const [varName, pos] of variantVars) {
        const inferred = resolved.get(varName) ?? null;
        if (!inferred) continue;
        const declared = declaredTypeMap.get(varName) ?? '';
        if (inferred.toLowerCase() === declared) continue; // 宣言型と同じは冗長
        hints.push({ varName, inferredType: inferred, ...pos, kind: 'var' });
    }
    return hints;
}

/**
 * 手続きのパラメーター・戻り型・変数の全ヒントを返す。
 * `inferVariantTypes`（変数のみ）の上位 API。
 */
export function inferProcedureHints(
    proc: ProcedureDeclaration,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
): VarHint[] {
    const hints: VarHint[] = [];

    // 変数宣言ヒント
    hints.push(...inferVariantTypes(proc, allProcs, memo));

    // パラメーターヒント（型なし / Variant）
    hints.push(...collectParameterHints(proc));

    // 戻り型ヒント（Function で型なし / Variant の場合）
    const ret = inferReturnTypeHint(proc, allProcs, memo);
    if (ret) hints.push(ret);

    return hints;
}

// ─── 代入収集（再帰的にボディを走査） ────────────────────────────────────────

function collectAssignmentTypes(
    stmts: Statement[],
    targets: Map<string, { line: number; column: number; endColumn: number }>,
    resolved: Map<string, InferredType>,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): void {
    for (const stmt of stmts) {
        // 代入文: x = expr / Set x = expr
        if (stmt.type === 'AssignmentStatement' || stmt.type === 'SetStatement') {
            const s = stmt as any;
            const lhs: any = s.left;
            if (lhs?.type === 'Identifier') {
                const name = (lhs as Identifier).name.toLowerCase();
                if (targets.has(name) && !resolved.has(name)) {
                    const t = inferExprType(s.right, allProcs, memo, depth);
                    if (t) resolved.set(name, t);
                }
            }
        }

        // For Each variable: For Each varName In collection
        if (stmt.type === 'ForEachStatement') {
            const fe = stmt as any;
            const varName = fe.variable?.name?.toLowerCase?.();
            if (varName && targets.has(varName) && !resolved.has(varName)) {
                const t = inferCollectionElementType(fe.collection, allProcs, memo, depth);
                if (t) resolved.set(varName, t);
            }
        }

        // 再帰: if/for/while/with 等のボディも走査
        recurseBody(stmt, targets, resolved, allProcs, memo, depth);
    }
}

function recurseBody(
    stmt: Statement,
    targets: Map<string, { line: number; column: number; endColumn: number }>,
    resolved: Map<string, InferredType>,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): void {
    const s = stmt as any;
    const bodies: Statement[][] = [];
    if (Array.isArray(s.body))       bodies.push(s.body);
    if (Array.isArray(s.consequent)) bodies.push(s.consequent);
    if (Array.isArray(s.alternate))  bodies.push(s.alternate);
    if (Array.isArray(s.elseBody))   bodies.push(s.elseBody);
    if (s.cases) for (const c of s.cases) if (Array.isArray(c.body)) bodies.push(c.body);
    for (const body of bodies) {
        collectAssignmentTypes(body, targets, resolved, allProcs, memo, depth);
    }
}

// ─── パラメーターヒント収集 ───────────────────────────────────────────────────

/** 型なし / Variant パラメーターに "As Variant" ヒントを付ける */
function collectParameterHints(proc: ProcedureDeclaration): VarHint[] {
    const hints: VarHint[] = [];
    for (const param of proc.parameters) {
        if (param.isParamArray) continue; // ParamArray はスキップ
        const t = param.paramType?.toLowerCase();
        if (t) continue; // 型あり（As Variant を含む）はスキップ — 型なしのみヒント対象

        const loc = (param as any).loc;
        if (!loc) continue;

        hints.push({
            varName:       param.name,
            inferredType:  'Variant',
            line:          loc.start.line,
            column:        loc.start.column,
            endColumn:     loc.start.column + param.name.length,
            kind:          'param',
        });
    }
    return hints;
}

// ─── 戻り型ヒント（Function の場合のみ） ─────────────────────────────────────

/** 型なし / Variant 関数の戻り型を推論してヒントを返す */
function inferReturnTypeHint(
    proc: ProcedureDeclaration,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
): VarHint | null {
    if (!proc.isFunction && !proc.isProperty) return null;
    const declaredType = proc.returnType?.toLowerCase();
    if (declaredType) return null; // 型が明示されていれば（As Variant を含む）スキップ

    const inferred = inferFunctionReturnType(proc, allProcs, memo, 0);
    if (!inferred) return null;

    // パラメーターリストの ')' の直後にヒントを表示（例: GetLabel() As String）
    const nameLoc = (proc.name as any).loc;
    if (!nameLoc) return null;

    const endColumn = (proc as any).paramsEndColumn
        ?? nameLoc.end?.column
        ?? nameLoc.start.column + proc.name.name.length;

    return {
        varName:      proc.name.name,
        inferredType: inferred,
        line:         nameLoc.start.line,
        column:       nameLoc.start.column,
        endColumn,
        kind:         'return',
    };
}

// ─── コレクション要素型の推論（For Each 用） ─────────────────────────────────

function inferCollectionElementType(
    collection: Expression,
    _allProcs: Map<string, ProcedureDeclaration>,
    _memo: Map<string, InferredType>,
    _depth: number,
): InferredType {
    if (!collection) return null;
    // Split() → 要素は String
    if (collection.type === 'CallExpression') {
        const callee = (collection as CallExpression).callee;
        if (callee.type === 'Identifier') {
            const name = (callee as Identifier).name.toLowerCase();
            if (name === 'split') return 'String';
        }
    }
    // .Keys / .Values / .Items → Dictionary 系は Variant のまま
    return null;
}

// ─── 式の型推論 ───────────────────────────────────────────────────────────────

function inferExprType(
    expr: Expression,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): InferredType {
    if (!expr || depth > MAX_DEPTH) return null;

    switch (expr.type) {
        case 'NumberLiteral': {
            const val = (expr as NumberLiteral).value;
            return Number.isInteger(val) ? 'Long' : 'Double';
        }
        case 'StringLiteral': return 'String';
        case 'DateLiteral':   return 'Date';

        case 'Identifier': {
            const name = (expr as Identifier).name.toLowerCase();
            if (name === 'true' || name === 'false') return 'Boolean';
            if (name === 'nothing')                  return 'Object';
            if (name === 'null')                     return 'Null';
            if (name === 'empty')                    return 'Empty';
            return null;
        }

        case 'BinaryExpression': {
            const be = expr as any;
            if (be.operator === '&') return 'String'; // 文字列連結
            const l = inferExprType(be.left,  allProcs, memo, depth);
            const r = inferExprType(be.right, allProcs, memo, depth);
            if (l === 'Double' || r === 'Double') return 'Double';
            if (l === 'Long'   && r === 'Long')   return 'Long';
            if (l === 'String' && r === 'String') return 'String';
            return null;
        }

        case 'CallExpression': {
            const callee = (expr as CallExpression).callee;
            if (callee.type === 'Identifier') {
                const name = (callee as Identifier).name;
                if (name.toLowerCase() === 'createobject') {
                    const ce = expr as CallExpression;
                    if (ce.args.length > 0 && ce.args[0].type === 'StringLiteral') {
                        return resolveProgIdType((ce.args[0] as StringLiteral).value);
                    }
                    return 'Object';
                }
                const builtin = Evaluator.BUILTIN_RETURN_TYPES[name.toLowerCase()];
                if (builtin) return builtin;
                return lookupReturnType(name, allProcs, memo, depth);
            }
            return null;
        }

        case 'ParenthesizedExpression':
            return inferExprType((expr as any).expression, allProcs, memo, depth);

        case 'UnaryExpression': {
            const ue = expr as any;
            if (ue.operator === 'Not') return 'Boolean';
            return inferExprType(ue.argument, allProcs, memo, depth);
        }

        default:
            return null;
    }
}

// ─── 関数の戻り型解決（メモ化・再帰） ─────────────────────────────────────────

function lookupReturnType(
    funcName: string,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): InferredType {
    const key = funcName.toLowerCase();
    if (memo.has(key)) return memo.get(key)!;

    const proc = allProcs.get(key);
    if (!proc || !proc.isFunction) { memo.set(key, null); return null; }

    // 明示的な非 Variant 戻り型があればそれを使う
    if (proc.returnType) {
        const rt = proc.returnType.toLowerCase();
        if (rt !== 'variant') {
            const result = normalizeType(proc.returnType);
            memo.set(key, result);
            return result;
        }
    }

    // Variant / 未宣言 → 関数ボディを再帰推論
    memo.set(key, null); // 無限再帰防止センチネル
    const inferred = inferFunctionReturnType(proc, allProcs, memo, depth + 1);
    memo.set(key, inferred);
    return inferred;
}

/** 関数ボディ内の「関数名 = 式」代入から戻り型を推論 */
function inferFunctionReturnType(
    proc: ProcedureDeclaration,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): InferredType {
    if (depth > MAX_DEPTH) return null;
    const funcName = proc.name.name.toLowerCase();
    return searchReturnAssignment(proc.body, funcName, allProcs, memo, depth);
}

function searchReturnAssignment(
    stmts: Statement[],
    funcName: string,
    allProcs: Map<string, ProcedureDeclaration>,
    memo: Map<string, InferredType>,
    depth: number,
): InferredType {
    for (const stmt of stmts) {
        if (stmt.type === 'AssignmentStatement') {
            const s = stmt as any;
            if (s.left?.type === 'Identifier' && s.left.name.toLowerCase() === funcName) {
                const t = inferExprType(s.right, allProcs, memo, depth);
                if (t) return t;
            }
        }
        // 再帰走査
        const s = stmt as any;
        for (const key of ['body', 'consequent', 'alternate', 'elseBody']) {
            if (Array.isArray(s[key])) {
                const found = searchReturnAssignment(s[key], funcName, allProcs, memo, depth);
                if (found) return found;
            }
        }
        if (s.cases) {
            for (const c of s.cases) {
                const found = searchReturnAssignment(c.body ?? [], funcName, allProcs, memo, depth);
                if (found) return found;
            }
        }
    }
    return null;
}

// ─── ユーティリティ ───────────────────────────────────────────────────────────

function resolveProgIdType(progId: string): string {
    const lower = progId.toLowerCase();
    if (lower.includes('dictionary'))  return 'Dictionary';
    if (lower.includes('filesystem'))  return 'FileSystemObject';
    if (lower.includes('range'))       return 'Range';
    if (lower.includes('worksheet'))   return 'Worksheet';
    if (lower.includes('workbook'))    return 'Workbook';
    return progId; // 未知 ProgID はそのまま返す（"As Excel.Application" 等で表示）
}

function normalizeType(t: string): string {
    const map: Record<string, string> = {
        long: 'Long', integer: 'Integer', string: 'String',
        boolean: 'Boolean', double: 'Double', single: 'Single',
        date: 'Date', object: 'Object', byte: 'Byte',
        currency: 'Currency', variant: 'Variant', decimal: 'Decimal',
    };
    return map[t.toLowerCase()] ?? t;
}

/**
 * ローカル変数の `As Object` 宣言について、その変数が宣言されているプロシージャ内の
 * `Set varName = CreateObject(...)` 代入から型を推論する。
 */
export function inferLocalVarType(varName: string, procName: string | null, statements: Statement[]): InferredType {
    if (!procName) return null;
    const lower = varName.toLowerCase();
    const procLower = procName.toLowerCase();
    for (const stmt of statements) {
        if (stmt.type !== 'ProcedureDeclaration') continue;
        const proc = stmt as ProcedureDeclaration;
        if (proc.name.name.toLowerCase() !== procLower) continue;
        return searchCreateObjectAssignment(proc.body, lower);
    }
    return null;
}

/**
 * モジュールレベルの `As Object` 変数について、全プロシージャボディを走査して
 * `Set varName = CreateObject(...)` 代入から型を推論する。
 */
export function inferModuleVarType(varName: string, statements: Statement[]): InferredType {
    const lower = varName.toLowerCase();
    for (const stmt of statements) {
        if (stmt.type !== 'ProcedureDeclaration') continue;
        const proc = stmt as ProcedureDeclaration;
        const t = searchCreateObjectAssignment(proc.body, lower);
        if (t) return t;
    }
    return null;
}

function collectCreateObjectProgIds(stmts: Statement[], varName: string): string[] {
    const results: string[] = [];
    for (const stmt of stmts) {
        if (stmt.type === 'SetStatement') {
            const s = stmt as any;
            if (s.left?.type === 'Identifier' && s.left.name.toLowerCase() === varName) {
                if (s.right?.type === 'CallExpression') {
                    const callee = s.right.callee;
                    if (callee?.type === 'Identifier' && callee.name.toLowerCase() === 'createobject') {
                        const args = s.right.args ?? [];
                        const t = args.length > 0 && args[0].type === 'StringLiteral'
                            ? resolveProgIdType(args[0].value)
                            : 'Object';
                        if (t) results.push(t);
                    }
                }
            }
        }
        const s = stmt as any;
        for (const key of ['body', 'consequent', 'alternate', 'elseBody']) {
            if (Array.isArray(s[key])) results.push(...collectCreateObjectProgIds(s[key], varName));
        }
        if (s.cases) {
            for (const c of s.cases) results.push(...collectCreateObjectProgIds(c.body ?? [], varName));
        }
    }
    return results;
}

function searchCreateObjectAssignment(stmts: Statement[], varName: string): InferredType {
    const types = collectCreateObjectProgIds(stmts, varName);
    if (types.length === 0) return null;
    // 複数の異なる型が存在する場合は曖昧なので null
    const uniq = new Set(types);
    return uniq.size === 1 ? types[0] : null;
}

/** AST のトップレベル手続きリストから名前→宣言のマップを作る */
export function buildProcMap(stmts: Statement[]): Map<string, ProcedureDeclaration> {
    const map = new Map<string, ProcedureDeclaration>();
    for (const stmt of stmts) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            map.set(proc.name.name.toLowerCase(), proc);
        } else if (stmt.type === 'ClassDeclaration') {
            for (const proc of (stmt as any).procedures ?? []) {
                map.set((proc as ProcedureDeclaration).name.name.toLowerCase(), proc);
            }
        }
    }
    return map;
}

/** カーソル行（0-based）を含む手続きを返す */
export function findProcAtLine(
    stmts: Statement[],
    line: number,
): ProcedureDeclaration | null {
    for (const stmt of stmts) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            const loc = (proc as any).loc;
            if (loc && loc.start.line - 1 <= line && line <= loc.end.line - 1) {
                return proc;
            }
        } else if (stmt.type === 'ClassDeclaration') {
            for (const proc of (stmt as any).procedures ?? []) {
                const loc = (proc as any).loc;
                if (loc && loc.start.line - 1 <= line && line <= loc.end.line - 1) {
                    return proc as ProcedureDeclaration;
                }
            }
        }
    }
    return null;
}
