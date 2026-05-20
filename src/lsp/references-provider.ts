import { Statement } from '../engine/parser';
import {
    buildScopedSymbolTable,
    findEnclosingScope,
    getWordAtPosition,
    lookupSymbol,
    ProcedureScope,
} from './symbol-table';

export interface LocationInfo {
    uri: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class ReferencesProvider {
    private uri: string = '';

    setDocumentUri(uri: string): void {
        this.uri = uri;
    }

    getReferences(
        statements: Statement[],
        sourceText: string,
        line: number,
        character: number,
        includeDeclaration: boolean,
    ): LocationInfo[] {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return [];
        return findAllReferences(sourceText, word, this.uri, statements, includeDeclaration, line);
    }
}

export function findAllReferences(
    sourceText: string,
    targetWord: string,
    uri: string,
    statements: Statement[],
    includeDeclaration: boolean,
    cursorLine?: number,
): LocationInfo[] {
    const refs: LocationInfo[] = [];
    const targetLower = targetWord.toLowerCase();

    const table = buildScopedSymbolTable(statements);

    // Determine scope of the symbol: local (inside a procedure) or module-level
    const resolvedLine = cursorLine ?? 0;
    const declEntry = lookupSymbol(targetWord, resolvedLine, table);
    const enclosingScope = findEnclosingScope(table.procedures, resolvedLine);

    // If the symbol is declared locally, find which procedure owns it
    let ownerScope: ProcedureScope | null = null;
    if (enclosingScope) {
        const local = enclosingScope.localSymbols.get(targetLower);
        if (local) ownerScope = enclosingScope;
    }
    // Fall back: check all procedure scopes for a local that matches
    if (!ownerScope) {
        for (const scope of table.procedures) {
            if (scope.localSymbols.has(targetLower)) {
                ownerScope = scope;
                break;
            }
        }
    }

    const lines = sourceText.split('\n');
    const pattern = new RegExp(`(?<![a-zA-Z0-9_])${escapeRegex(targetWord)}(?![a-zA-Z0-9_])`, 'gi');

    for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
        // If this is a local symbol, restrict search to the owning procedure's line range
        if (ownerScope && (lineIdx < ownerScope.range.start.line || lineIdx > ownerScope.range.end.line)) {
            continue;
        }

        const lineText = lines[lineIdx];

        const trimmed = lineText.trimStart();
        if (trimmed.startsWith("'") || trimmed.toLowerCase().startsWith('rem ')) continue;

        let match: RegExpExecArray | null;
        pattern.lastIndex = 0;

        while ((match = pattern.exec(lineText)) !== null) {
            const charIdx = match.index;

            if (isInsideString(lineText, charIdx)) continue;
            if (isAfterComment(lineText, charIdx)) continue;

            const range = {
                start: { line: lineIdx, character: charIdx },
                end: { line: lineIdx, character: charIdx + targetWord.length },
            };

            if (!includeDeclaration && declEntry) {
                const d = declEntry.range.start;
                if (lineIdx === d.line && charIdx === d.character) continue;
            }

            refs.push({ uri, range });
        }
    }

    return refs;
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isInsideString(lineText: string, pos: number): boolean {
    let inString = false;
    for (let i = 0; i < pos; i++) {
        if (lineText[i] === '"') {
            if (i + 1 < lineText.length && lineText[i + 1] === '"') {
                i++;
            } else {
                inString = !inString;
            }
        }
    }
    return inString;
}

function isAfterComment(lineText: string, pos: number): boolean {
    let inString = false;
    for (let i = 0; i < pos; i++) {
        if (lineText[i] === '"') {
            if (i + 1 < lineText.length && lineText[i + 1] === '"') {
                i++;
            } else {
                inString = !inString;
            }
        } else if (lineText[i] === "'" && !inString) {
            return true;
        }
    }
    return false;
}
