import { Statement } from '../engine/parser';
import { buildSymbolTable, getWordAtPosition } from './symbol-table';

export interface Hover {
    contents: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class HoverProvider {
    getHoverInfo(statements: Statement[], sourceText: string, line: number, character: number): Hover | null {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return null;

        const symbols = buildSymbolTable(statements);
        const entry = symbols.get(word.toLowerCase());
        if (!entry) return null;

        return {
            contents: entry.displayText,
            range: entry.range,
        };
    }
}
