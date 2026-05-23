import { Statement } from '../engine/parser';

export interface FoldingRange {
    startLine: number; // 0-based
    endLine: number;   // 0-based
}

// AST node types that represent foldable blocks
const BLOCK_TYPES = new Set([
    'ProcedureDeclaration',
    'ForStatement',
    'ForEachStatement',
    'DoWhileStatement',
    'WhileStatement',
    'WithStatement',
    'IfStatement',
    'SelectCaseStatement',
    'TypeDeclaration',
    'EnumDeclaration',
    'ClassDeclaration',
]);

export class FoldingRangeProvider {
    getFoldingRanges(statements: Statement[]): FoldingRange[] {
        const ranges: FoldingRange[] = [];
        this.collectRanges(statements, ranges);
        return ranges;
    }

    private collectRanges(statements: Statement[], ranges: FoldingRange[]) {
        for (const stmt of statements) {
            this.collectFromStatement(stmt, ranges);
        }
    }

    private collectFromStatement(stmt: Statement, ranges: FoldingRange[]) {
        if (!stmt.loc) return;

        if (BLOCK_TYPES.has(stmt.type)) {
            const start = stmt.loc.start.line - 1; // parser is 1-based, VS Code is 0-based
            const end   = stmt.loc.end.line - 1;
            if (end > start) {
                ranges.push({ startLine: start, endLine: end });
            }
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const s = stmt as any;

        switch (stmt.type) {
            case 'ProcedureDeclaration':
            case 'ForStatement':
            case 'ForEachStatement':
            case 'DoWhileStatement':
            case 'WhileStatement':
            case 'WithStatement':
                this.collectRanges(s.body ?? [], ranges);
                break;

            case 'IfStatement':
                this.collectRanges(s.consequent ?? [], ranges);
                if (Array.isArray(s.alternate)) {
                    this.collectRanges(s.alternate, ranges);
                } else if (s.alternate) {
                    // ElseIf chain: nested IfStatement
                    this.collectFromStatement(s.alternate, ranges);
                }
                break;

            case 'SelectCaseStatement':
                for (const c of s.cases ?? []) {
                    this.collectRanges(c.body ?? [], ranges);
                }
                this.collectRanges(s.elseBody ?? [], ranges);
                break;

            case 'ClassDeclaration':
                this.collectRanges(s.procedures ?? [], ranges);
                this.collectRanges(s.body ?? [], ranges);
                break;
        }
    }
}
