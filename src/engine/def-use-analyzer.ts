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
 * 実装: CFG + 到達定義（Phase 4B）+ 生変数（Phase 4C）による精密解析。
 *   - 到達定義: 使用点に到達する定義集合から「範囲外の定義」を検出 → inputs
 *   - 生変数: 範囲内の定義変数が範囲後で生きているか判定 → outputs
 *   - 条件分岐・ループを含む場合も正確に分類できる。
 */

import { ProcedureDeclaration } from './parser';
import { buildCFG, CFG } from './cfg';
import {
    computeReachingDefs,
    buildDefUseChains,
    usePointKey,
    UsePoint,
} from './reaching-defs';
import { computeLiveVars, LiveVarsResult, getStmtDefs, getStmtUses } from './live-vars';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface DefUseResult {
    /** 範囲外で定義され、範囲内で使用される変数（関数の ByVal 引数候補） */
    inputs: string[];
    /** 範囲内で定義され、範囲外（後）で使用される変数（ByRef 引数候補） */
    outputs: string[];
    /** 範囲内だけで完結する変数（抽出後のローカル変数候補） */
    locals: string[];
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
    const cfg = buildCFG(proc);

    // 到達定義解析（パラメーターを暗黙 def として登録）
    const paramNames = new Set(proc.parameters.map(p => p.name.toLowerCase()));
    const rdResult = computeReachingDefs(cfg, paramNames);

    // 生変数解析（ByRef 引数・関数戻り値を出口で常に生きているとみなす）
    const alwaysLive = new Set<string>();
    if (proc.isFunction || proc.isProperty) alwaysLive.add(proc.name.name.toLowerCase());
    for (const param of proc.parameters) {
        if (!param.isByVal) alwaysLive.add(param.name.toLowerCase());
    }
    const lvResult = computeLiveVars(cfg, alwaysLive);

    // プロシージャ内で少なくとも 1 回 def される変数名 + パラメーター
    // （VBA 組み込み関数名などの偽変数をフィルタリングするために使用）
    const knownVars = new Set<string>(rdResult.allDefs.map(d => d.varName));
    for (const p of paramNames) knownVars.add(p);

    // ─── 範囲内ステートメントの収集 ─────────────────────────────────────────
    const inRangeUsePoints: UsePoint[] = [];
    const allDefsInRange = new Set<string>();

    for (const block of cfg.blocks) {
        if (block.kind !== 'normal') continue;
        for (let i = 0; i < block.stmts.length; i++) {
            const stmt = block.stmts[i];
            const line = (stmt as any).loc?.start.line ?? 0;
            if (line < startLine || line > endLine) continue;

            for (const v of getStmtDefs(stmt)) allDefsInRange.add(v);
            for (const v of getStmtUses(stmt)) {
                if (knownVars.has(v)) {
                    inRangeUsePoints.push({ blockId: block.id, stmtIdx: i, varName: v });
                }
            }
        }
    }

    // ─── inputs: 範囲外の定義が到達する使用点 ─────────────────────────────────
    const chains = buildDefUseChains(cfg, rdResult, inRangeUsePoints);
    const inputs = new Set<string>();

    for (const up of inRangeUsePoints) {
        const reachDefs = chains.get(usePointKey(up)) ?? new Set();
        const hasOutsideDef = [...reachDefs].some(d =>
            d.stmtIdx === -1          || // パラメーター（暗黙 def）
            d.line < startLine        ||
            d.line > endLine,
        );
        // 「本物の先行定義」がない = inputs とみなすケース:
        //   1. 到達定義が空 → VBA 暗黙初期化値に依存
        //   2. すべての到達定義が「同ブロック・同位置以降」= ループバックエッジ経由の
        //      前イテレーションの残像であり、このパス上に真の先行定義がない
        const noGenuinePriorDef = reachDefs.size === 0 || [...reachDefs].every(d =>
            d.blockId === up.blockId && d.stmtIdx >= up.stmtIdx,
        );

        if (hasOutsideDef || noGenuinePriorDef) inputs.add(up.varName);
    }

    // ─── outputs: 範囲内の定義変数が endLine より後で生きている ───────────────
    const outputs = new Set<string>();
    for (const varName of allDefsInRange) {
        if (isLiveAfterEndLine(varName, cfg, lvResult, endLine)) outputs.add(varName);
    }

    // ─── locals: 範囲内で完結（inputs でも outputs でもない） ─────────────────
    const locals: string[] = [];
    for (const varName of allDefsInRange) {
        if (!inputs.has(varName) && !outputs.has(varName)) locals.push(varName);
    }

    return {
        inputs:  [...inputs].sort(),
        outputs: [...outputs].sort(),
        locals:  locals.sort(),
    };
}

// ─── 内部ユーティリティ ───────────────────────────────────────────────────────

/**
 * 変数 varName が endLine より後の任意の時点で生きているかを判定する。
 *
 * 各ブロックについて blockOut から後ろ向きにシミュレーションし、
 * line > endLine のステートメント直前（= そのステートメントかそれ以降で使用）で
 * varName が生きていれば true を返す。
 */
function isLiveAfterEndLine(
    varName: string,
    cfg: CFG,
    lvResult: LiveVarsResult,
    endLine: number,
): boolean {
    for (const block of cfg.blocks) {
        if (block.kind !== 'normal') continue;

        const hasAfterLine = block.stmts.some(
            s => ((s as any).loc?.start.line ?? 0) > endLine,
        );
        if (!hasAfterLine) continue;

        // ブロック出口の生変数集合を出発点に後ろ向きシミュレーション
        const live = new Set<string>(lvResult.blockOut.get(block.id)!);

        for (let i = block.stmts.length - 1; i >= 0; i--) {
            const stmt = block.stmts[i];
            const line = (stmt as any).loc?.start.line ?? 0;

            // live_before(stmt i) = (live_after − defs) ∪ uses
            for (const v of getStmtDefs(stmt)) live.delete(v);
            for (const v of getStmtUses(stmt)) live.add(v);

            // live_before: stmt i またはそれ以降で varName が使われるか
            if (line > endLine && live.has(varName)) return true;
        }
    }
    return false;
}
