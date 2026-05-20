import { Statement } from '../engine/parser';
import { getWordAtPosition } from './symbol-table';
import { findAllReferences, LocationInfo } from './references-provider';

export interface TextEdit {
    range: LocationInfo['range'];
    newText: string;
}

export class RenameProvider {
    private uri: string = '';

    setDocumentUri(uri: string): void {
        this.uri = uri;
    }

    getRename(
        statements: Statement[],
        sourceText: string,
        line: number,
        character: number,
        newName: string,
    ): TextEdit[] | null {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return null;

        const refs = findAllReferences(sourceText, word, this.uri, statements, true, line);
        if (refs.length === 0) return null;

        return refs.map(ref => ({ range: ref.range, newText: newName }));
    }
}
