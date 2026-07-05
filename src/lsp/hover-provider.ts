import { Statement, ProcedureDeclaration } from '../engine/parser';
import { buildCFG, BasicBlock } from '../engine/cfg';
import { computeReachingDefs, buildDefUseChains, usePointKey } from '../engine/reaching-defs';
import {
    buildScopedSymbolTable,
    getWordAtPosition,
    lookupSymbolWithContext,
    SymbolKind,
    SymbolLookupResult,
} from './symbol-table';
import { inferModuleVarType } from './variant-type-inferencer';

export interface Hover {
    contents: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class HoverProvider {
    getHoverInfo(
        statements: Statement[],
        sourceText: string,
        line: number,
        character: number,
    ): Hover | null {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return null;

        const table = buildScopedSymbolTable(statements);
        const ctx = lookupSymbolWithContext(word, line, table);

        const reachingInfo = this.getReachingDefsInfo(statements, word, line);

        if (!ctx && !reachingInfo) return null;

        const parts: string[] = [];

        if (ctx) {
            let displayText = ctx.entry.displayText;
            // モジュール変数が As Object の場合、全プロシージャを走査して CreateObject 代入から型を推論
            if (ctx.entry.kind === 'module-var' && /\bAs Object\b/i.test(displayText)) {
                const inferred = inferModuleVarType(ctx.entry.name, statements);
                if (inferred) displayText = displayText.replace(/\bAs Object\b/i, `As ${inferred}`);
            }
            parts.push(`\`\`\`vba\n${displayText}\n\`\`\``);
            const ctxLine = buildContextLine(ctx);
            if (ctxLine) parts.push(ctxLine);
        }

        if (reachingInfo) parts.push(reachingInfo);

        const range = ctx?.entry.range ?? {
            start: { line, character },
            end: { line, character: character + word.length },
        };
        return { contents: parts.join('\n\n'), range };
    }

    private getReachingDefsInfo(
        statements: Statement[],
        word: string,
        line0: number,
    ): string | null {
        const line1 = line0 + 1;

        const proc = statements.find(s => {
            if (s.type !== 'ProcedureDeclaration') return false;
            const p = s as ProcedureDeclaration;
            return p.loc != null && p.loc.start.line <= line1 && p.loc.end.line >= line1;
        }) as ProcedureDeclaration | undefined;
        if (!proc) return null;

        const varLower    = word.toLowerCase();
        const paramNames  = new Set(proc.parameters.map(p => p.name.toLowerCase()));

        try {
            const cfg      = buildCFG(proc);
            const rdResult = computeReachingDefs(cfg, paramNames);

            // カーソル行のブロック・stmtIdx を探す
            let foundBlock: BasicBlock | null = null;
            let foundIdx = -1;
            for (const block of cfg.blocks) {
                if (block.kind !== 'normal') continue;
                for (let i = 0; i < block.stmts.length; i++) {
                    if (((block.stmts[i] as any).loc?.start.line ?? 0) === line1) {
                        foundBlock = block;
                        foundIdx   = i;
                        break;
                    }
                }
                if (foundBlock) break;
            }
            if (!foundBlock || foundIdx < 0) return null;

            const up     = { blockId: foundBlock.id, stmtIdx: foundIdx, varName: varLower };
            const chains = buildDefUseChains(cfg, rdResult, [up]);
            const defs   = chains.get(usePointKey(up));
            if (!defs || defs.size === 0) return null;

            const defList = [...defs].sort((a, b) => a.line - b.line);
            const desc    = defList.map(d => d.stmtIdx === -1 ? 'パラメーター' : `行 ${d.line}`).join(', ');
            return `到達定義: ${desc}`;
        } catch {
            return null;
        }
    }
}

// ─── context line builder ─────────────────────────────────────────────────────

function buildContextLine(ctx: SymbolLookupResult): string {
    const kindLabel = kindContextLabel(ctx.entry.kind, ctx.procName);
    // HoverProvider は常に現在ファイルの AST のみを検索するため "this module" で固定
    return [kindLabel, 'this module'].filter(Boolean).join(' · ');
}

function kindContextLabel(kind: SymbolKind, procName: string | null): string {
    switch (kind) {
        case 'module-var':  return 'Module variable';
        case 'local-var':   return procName ? `Local variable in \`${procName}\`` : 'Local variable';
        case 'param':       return procName ? `Parameter of \`${procName}\`` : 'Parameter';
        case 'for':
        case 'for-each':    return procName ? `Loop variable in \`${procName}\`` : 'Loop variable';
        case 'const':       return procName ? `Constant in \`${procName}\`` : 'Module constant';
        case 'event':       return 'Event';
        case 'class':       return 'Class';
        case 'procedure':   return '';
    }
}
