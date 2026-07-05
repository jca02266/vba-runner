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

    // カーソルあり（同ファイル検索）: カーソル位置のプロシージャ内でローカル宣言があれば
    // そのプロシージャに検索を限定する。
    // カーソルなし（クロスファイル検索）: ownerScope は設定しない。
    // 代わりに行ループ内でプロシージャ単位の shadowing を処理する。
    let ownerScope: ProcedureScope | null = null;
    if (cursorLine !== undefined && enclosingScope) {
        const local = enclosingScope.localSymbols.get(targetLower);
        if (local) ownerScope = enclosingScope;
    }

    const lines = sourceText.split('\n');
    const pattern = new RegExp(`(?<![a-zA-Z0-9_])${escapeRegex(targetWord)}(?![a-zA-Z0-9_])`, 'gi');

    for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
        if (ownerScope) {
            // ローカルシンボル: 所有プロシージャの範囲のみ
            if (lineIdx < ownerScope.range.start.line || lineIdx > ownerScope.range.end.line) continue;
        } else {
            // モジュールレベルシンボル（クロスファイル含む）:
            // 各行が属するプロシージャがそのシンボルをローカル宣言していれば
            // shadowing されているため除外する（プロシージャ単位）
            const scope = findEnclosingScope(table.procedures, lineIdx);
            if (scope && scope.localSymbols.has(targetLower)) continue;
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
