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
    /** ソース行（1-based、パーサーの loc 由来。0 は位置情報なし） */
    line: number;
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

    const { blockOut } = computeLiveVars(cfg, alwaysLive);

    const results: DeadStore[] = [];
    const seen = new Set<string>(); // 同一変数の重複報告を抑制

    for (const block of cfg.blocks) {
        if (block.kind !== 'normal') continue;

        // ブロック末尾の生変数集合を出発点に後ろ向きシミュレーション
        const live = new Set<string>(blockOut.get(block.id)!);

        for (let i = block.stmts.length - 1; i >= 0; i--) {
            const stmt = block.stmts[i];
            const defs = getStmtDefs(stmt);
            const uses = getStmtUses(stmt);

            for (const v of defs) {
                if (!live.has(v) && !paramNames.has(v)) {
                    const line = (stmt as any).loc?.start.line ?? 0;
                    const key = `${v}:${line}`;
                    if (!seen.has(key)) {
                        seen.add(key);
                        results.push({ varName: v, line });
                    }
                }
            }

            // 後ろ向き: live_before = (live_after − defs) ∪ uses
            for (const v of defs) live.delete(v);
            for (const v of uses) live.add(v);
        }
    }

    // 行番号昇順でソート
    return results.sort((a, b) => a.line - b.line);
}
