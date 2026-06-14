/**
 * デッドストア検出（Dead Store Detection）— Phase 4D
 *
 * 「代入されたが、次に使われる前に上書きされるか、プロシージャ終了まで参照されない変数」
 * を CFG + 生変数解析（Phase 4C）を使って検出する。
 *
 * 対象: ローカル変数への単純代入（Identifier lhs）のみ。
 *   除外: パラメーター（ByRef は呼び元への出力）、関数戻り値変数、配列要素・メンバー代入
 */

import { ProcedureDeclaration } from './parser';
import { buildCFG } from './cfg';
import { computeLiveVars, getStmtDefs, getStmtUses } from './live-vars';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface DeadStore {
    /** 代入先変数名（小文字正規化済み） */
    varName: string;
    /** ステートメント種別（AssignmentStatement / SetStatement / ForStatement 等） */
    stmtType: string;
    /** ソース行（1-based、0 は位置情報なし） */
    line: number;
    /** ソース列（1-based、0 は位置情報なし）— LHS 変数の先頭 */
    column: number;
    /** LHS 変数の末尾列（1-based） */
    endColumn: number;
}

// ─── 公開 API ─────────────────────────────────────────────────────────────────

/**
 * プロシージャ内のデッドストアを検出する。
 */
export function findDeadStores(proc: ProcedureDeclaration): DeadStore[] {
    const cfg = buildCFG(proc);

    // 出口で常に生きているとみなす変数
    const alwaysLive = new Set<string>();

    // Function/Property の戻り値変数（関数名への代入は呼び元に渡る）
    if (proc.isFunction || proc.isProperty) {
        alwaysLive.add(proc.name.name.toLowerCase());
    }

    // ByRef パラメーターは呼び元への出力として扱う
    const paramNames = new Set<string>();
    for (const param of proc.parameters) {
        const name = param.name.toLowerCase();
        paramNames.add(name);
        if (!param.isByVal) alwaysLive.add(name); // ByRef
    }

    // モジュールレベル変数（プロシージャ内でローカル宣言されていない変数）は
    // 他のプロシージャ/プロパティから読まれる可能性があるため常に生存扱いにする。
    // これにより、クラスモジュールのフィールドへの代入を誤ってデッドストアと判定しない。
    const localDecls = collectLocalDecls(proc.body);
    for (const block of cfg.blocks) {
        for (const stmt of block.stmts) {
            for (const v of getStmtDefs(stmt)) {
                if (!localDecls.has(v) && !paramNames.has(v)) {
                    alwaysLive.add(v);
                }
            }
        }
    }

    const { blockOut } = computeLiveVars(cfg, alwaysLive);

    const results: DeadStore[] = [];
    const seen = new Set<string>();

    for (const block of cfg.blocks) {
        if (block.kind !== 'normal') continue;

        const live = new Set<string>(blockOut.get(block.id)!);

        for (let i = block.stmts.length - 1; i >= 0; i--) {
            const stmt = block.stmts[i];
            const defs = getStmtDefs(stmt);
            const uses = getStmtUses(stmt);

            for (const v of defs) {
                if (!live.has(v) && !paramNames.has(v)) {
                    const loc = lhsLoc(stmt, v);
                    const key = `${v}:${loc.line}:${loc.column}`;
                    if (!seen.has(key)) {
                        seen.add(key);
                        results.push({
                            varName:   v,
                            stmtType:  stmt.type,
                            line:      loc.line,
                            column:    loc.column,
                            endColumn: loc.endColumn,
                        });
                    }
                }
            }

            // 後ろ向き: live_before = (live_after − defs) ∪ uses
            for (const v of defs) live.delete(v);
            for (const v of uses) live.add(v);
        }
    }

    return results.sort((a, b) => a.line - b.line || a.column - b.column);
}

// ─── 内部ユーティリティ ───────────────────────────────────────────────────────

/** プロシージャ本体内でローカル宣言された変数名（小文字）を再帰的に収集する */
function collectLocalDecls(stmts: any[]): Set<string> {
    const locals = new Set<string>();
    for (const stmt of stmts) {
        if (stmt.type === 'VariableDeclaration') {
            for (const decl of stmt.declarations ?? []) {
                if (decl.name?.name) locals.add(decl.name.name.toLowerCase());
            }
        }
        for (const key of ['body', 'consequent', 'alternate']) {
            if (Array.isArray(stmt[key])) {
                for (const n of collectLocalDecls(stmt[key])) locals.add(n);
            }
        }
        for (const clause of stmt.cases ?? []) {
            for (const n of collectLocalDecls(clause.body ?? [])) locals.add(n);
        }
        if (Array.isArray(stmt.elseBody)) {
            for (const n of collectLocalDecls(stmt.elseBody)) locals.add(n);
        }
    }
    return locals;
}

function lhsLoc(stmt: any, varName: string): { line: number; column: number; endColumn: number } {
    // AssignmentStatement / SetStatement: LHS Identifier の loc を優先
    if (stmt.type === 'AssignmentStatement' || stmt.type === 'SetStatement') {
        const lhsLoc = stmt.left?.loc;
        if (lhsLoc) {
            return {
                line:      lhsLoc.start.line,
                column:    lhsLoc.start.column,
                endColumn: lhsLoc.end?.column ?? lhsLoc.start.column + varName.length,
            };
        }
    }
    // ForStatement: ループカウンター識別子の loc
    if (stmt.type === 'ForStatement') {
        const idLoc = stmt.identifier?.loc;
        if (idLoc) {
            return {
                line:      idLoc.start.line,
                column:    idLoc.start.column,
                endColumn: idLoc.end?.column ?? idLoc.start.column + varName.length,
            };
        }
    }
    // ForEachStatement
    if (stmt.type === 'ForEachStatement') {
        const vLoc = stmt.variable?.loc;
        if (vLoc) {
            return {
                line:      vLoc.start.line,
                column:    vLoc.start.column,
                endColumn: vLoc.end?.column ?? vLoc.start.column + varName.length,
            };
        }
    }
    // フォールバック: ステートメント全体の loc
    const stmtLoc = stmt.loc;
    const line   = stmtLoc?.start.line   ?? 0;
    const column = stmtLoc?.start.column ?? 0;
    return { line, column, endColumn: column + varName.length };
}
