/**
 * 到達定義解析（Reaching Definitions Analysis）— Phase 4B
 *
 * CFG の各基本ブロックについて「その点に到達している変数定義」の集合を算出する
 * 前向きデータフロー解析。
 *
 * IN[B]  = ⋃ OUT[P]   （P: B の先行ブロック）
 * OUT[B] = gen[B] ∪ (IN[B] − kill[B])
 *
 * gen[B]:  B 内でその後に上書きされない定義（最後の定義のみ）
 * kill[B]: B 内で定義される変数の、他ブロックにある全定義
 *
 * 主な用途:
 *   - Def-Use チェーン: ある使用点に到達しうる定義を列挙
 *   - 「潜在的な未初期化使用」の検出
 *   - Extract Function 精度向上（def-use-analyzer.ts の線形スキャンを置き換え）
 */

import { CFG } from './cfg';
import { ProcedureDeclaration } from './parser';
import { buildCFG } from './cfg';
import { getStmtDefs } from './live-vars';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface Definition {
    /** 変数名（小文字正規化済み） */
    varName: string;
    /** 定義が存在するブロック ID */
    blockId: number;
    /** ブロック内ステートメントインデックス */
    stmtIdx: number;
    /** ソース行（0 = 位置情報なし） */
    line: number;
}

export interface ReachingDefsResult {
    /** ブロック ID → そのブロック先頭に到達している定義集合 */
    blockIn:  Map<number, ReadonlySet<Definition>>;
    /** ブロック ID → そのブロック末尾から出ていく定義集合 */
    blockOut: Map<number, ReadonlySet<Definition>>;
    /** プログラム全体の全定義（デバッグ・追加解析用） */
    allDefs: readonly Definition[];
}

/** ある使用点（blockId, stmtIdx）に到達しうる定義の集合 */
export interface UsePoint {
    blockId: number;
    stmtIdx: number;
    varName: string;
}

// ─── 公開 API ─────────────────────────────────────────────────────────────────

/**
 * CFG 全体に対して到達定義解析を実行する。
 * @param paramNames プロシージャパラメーター名。entry ブロックの implicit def として扱う。
 */
export function computeReachingDefs(
    cfg: CFG,
    paramNames: ReadonlySet<string> = new Set(),
): ReachingDefsResult {
    // ─── Step 1: 全定義を収集 ────────────────────────────────────────────────
    const allDefs: Definition[] = [];
    const defById = new Map<string, Definition>();

    // パラメーターは entry ブロックへの暗黙的な定義として登録（stmtIdx = -1）
    for (const varName of paramNames) {
        const d: Definition = { varName, blockId: cfg.entry.id, stmtIdx: -1, line: 0 };
        const key = defKey(d);
        allDefs.push(d);
        defById.set(key, d);
    }

    for (const block of cfg.blocks) {
        for (let i = 0; i < block.stmts.length; i++) {
            const stmt = block.stmts[i];
            for (const varName of getStmtDefs(stmt)) {
                const line = (stmt as any).loc?.start.line ?? 0;
                const d: Definition = { varName, blockId: block.id, stmtIdx: i, line };
                const key = defKey(d);
                allDefs.push(d);
                defById.set(key, d);
            }
        }
    }

    // ─── Step 2: gen[B] / kill[B] を各ブロックで計算 ────────────────────────
    const genMap  = new Map<number, Set<string>>();  // blockId → Set<defKey>
    const killMap = new Map<number, Set<string>>();  // blockId → Set<defKey>

    for (const block of cfg.blocks) {
        // gen[B]: 各変数の最後の定義のみが外に出る
        const lastDefIdx = new Map<string, number>(); // varName → stmtIdx
        if (block.id === cfg.entry.id) {
            for (const v of paramNames) lastDefIdx.set(v, -1);
        }
        for (let i = 0; i < block.stmts.length; i++) {
            for (const v of getStmtDefs(block.stmts[i])) {
                lastDefIdx.set(v, i);
            }
        }

        const gen = new Set<string>();
        for (const [v, idx] of lastDefIdx) {
            gen.add(defKey({ varName: v, blockId: block.id, stmtIdx: idx }));
        }
        genMap.set(block.id, gen);

        // kill[B]: このブロックで定義される変数の、他ブロックにある全定義
        const defdHere = new Set(lastDefIdx.keys());
        const kill = new Set<string>();
        for (const d of allDefs) {
            if (d.blockId !== block.id && defdHere.has(d.varName)) {
                kill.add(defKey(d));
            }
        }
        killMap.set(block.id, kill);
    }

    // ─── Step 3: 前向きワークリスト ──────────────────────────────────────────
    const blockInKeys  = new Map<number, Set<string>>();
    const blockOutKeys = new Map<number, Set<string>>();
    for (const block of cfg.blocks) {
        blockInKeys.set(block.id, new Set());
        // entry は gen[entry] で初期化（パラメーター定義を含む）
        blockOutKeys.set(block.id, new Set(genMap.get(block.id)!));
    }

    const worklist = new Set<number>(cfg.blocks.map(b => b.id));
    while (worklist.size > 0) {
        const bid = worklist.values().next().value as number;
        worklist.delete(bid);
        const block = cfg.blocks.find(b => b.id === bid)!;

        // IN[B] = ⋃ OUT[P]
        const newIn = new Set<string>();
        for (const pred of block.preds) {
            for (const k of blockOutKeys.get(pred.id)!) newIn.add(k);
        }
        blockInKeys.set(block.id, newIn);

        // OUT[B] = gen[B] ∪ (IN[B] − kill[B])
        const gen  = genMap.get(block.id)!;
        const kill = killMap.get(block.id)!;
        const newOut = new Set<string>(gen);
        for (const k of newIn) {
            if (!kill.has(k)) newOut.add(k);
        }

        const oldOut = blockOutKeys.get(block.id)!;
        if (!strSetsEqual(oldOut, newOut)) {
            blockOutKeys.set(block.id, newOut);
            for (const succ of block.succs) worklist.add(succ.id);
        }
    }

    // ─── Step 4: キーを Definition オブジェクトに変換 ────────────────────────
    const blockIn  = new Map<number, ReadonlySet<Definition>>();
    const blockOut = new Map<number, ReadonlySet<Definition>>();
    for (const block of cfg.blocks) {
        blockIn.set(block.id,  keysToDefSet(blockInKeys.get(block.id)!,  defById));
        blockOut.set(block.id, keysToDefSet(blockOutKeys.get(block.id)!, defById));
    }

    return { blockIn, blockOut, allDefs };
}

/**
 * 到達定義結果から Def-Use チェーンを構築する。
 * @returns 使用点 (blockId, stmtIdx, varName) → 到達しうる定義集合
 */
export function buildDefUseChains(
    cfg: CFG,
    result: ReachingDefsResult,
    usePoints: readonly UsePoint[],
): Map<string, ReadonlySet<Definition>> {
    const chains = new Map<string, ReadonlySet<Definition>>();

    for (const up of usePoints) {
        const block = cfg.blocks.find(b => b.id === up.blockId);
        if (!block) continue;

        // ブロック先頭の到達定義から出発し、stmtIdx 直前まで前向きにシミュレーション
        // （ブロック内で up.varName が再定義された場合は最後の再定義のみが到達）
        const reachable = new Set<Definition>(
            [...result.blockIn.get(up.blockId)!].filter(d => d.varName === up.varName),
        );

        for (let i = 0; i < up.stmtIdx; i++) {
            const defs = getStmtDefs(block.stmts[i]).filter(v => v === up.varName);
            if (defs.length > 0) {
                // この変数は i で再定義された → それ以前の定義はすべてキル
                reachable.clear();
                const newDef = result.allDefs.find(
                    d => d.blockId === up.blockId && d.stmtIdx === i && d.varName === up.varName,
                );
                if (newDef) reachable.add(newDef);
            }
        }

        const key = usePointKey(up);
        chains.set(key, reachable);
    }

    return chains;
}

/** UsePoint の文字列キー */
export function usePointKey(up: UsePoint): string {
    return `${up.blockId}:${up.stmtIdx}:${up.varName}`;
}

// ─── プロシージャ全体に対する便利 API ────────────────────────────────────────

/** ProcedureDeclaration から直接 CFG を構築して到達定義解析を実行する */
export function analyzeReachingDefs(proc: ProcedureDeclaration): {
    cfg: CFG;
    result: ReachingDefsResult;
} {
    const cfg = buildCFG(proc);
    const paramNames = new Set(proc.parameters.map(p => p.name.toLowerCase()));
    const result = computeReachingDefs(cfg, paramNames);
    return { cfg, result };
}

// ─── 内部ユーティリティ ───────────────────────────────────────────────────────

function defKey(d: Pick<Definition, 'blockId' | 'stmtIdx' | 'varName'>): string {
    return `${d.blockId}:${d.stmtIdx}:${d.varName}`;
}

function keysToDefSet(keys: Set<string>, byId: Map<string, Definition>): ReadonlySet<Definition> {
    const s = new Set<Definition>();
    for (const k of keys) {
        const d = byId.get(k);
        if (d) s.add(d);
    }
    return s;
}

function strSetsEqual(a: Set<string>, b: Set<string>): boolean {
    if (a.size !== b.size) return false;
    for (const v of a) if (!b.has(v)) return false;
    return true;
}
