import { Statement } from '../engine/parser';
import { buildSymbolTable, getWordAtPosition } from './symbol-table';

export interface LocationInfo {
    uri: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class DefinitionProvider {
    private uri: string = '';

    setDocumentUri(uri: string): void {
        this.uri = uri;
    }

    getDefinition(statements: Statement[], sourceText: string, line: number, character: number): LocationInfo | null {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return null;

        const symbols = buildSymbolTable(statements);
        const entry = symbols.get(word.toLowerCase());
        if (!entry) return null;

        return { uri: this.uri, range: entry.range };
    }
}
