import {
    ClassDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    EventDeclaration,
    Statement,
} from '../compiler/parser';

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

    getDefinition(statements: Statement[], line: number, character: number): LocationInfo | null {
        return this.findDefinitionInStatements(statements, line, character);
    }

    private findDefinitionInStatements(statements: Statement[], line: number, character: number): LocationInfo | null {
        for (const stmt of statements) {
            const result = this.findDefinitionInStatement(stmt, line, character);
            if (result) return result;
        }
        return null;
    }

    private findDefinitionInStatement(stmt: Statement, line: number, character: number): LocationInfo | null {
        if (!stmt.loc) return null;

        // Check if position is within this statement's range
        const stmtLine = stmt.loc.start.line - 1;
        const stmtStartChar = stmt.loc.start.column - 1;
        const stmtEndLine = stmt.loc.end.line - 1;
        const stmtEndChar = stmt.loc.end.column - 1;

        if (line < stmtLine || line > stmtEndLine) {
            return null;
        }

        if (line === stmtLine && character < stmtStartChar) {
            return null;
        }

        if (line === stmtEndLine && character >= stmtEndChar) {
            return null;
        }

        // Check the specific statement type
        if (stmt.type === 'ProcedureDeclaration') {
            return this.definitionForProcedure(stmt as ProcedureDeclaration, line, character);
        } else if (stmt.type === 'ClassDeclaration') {
            return this.definitionForClass(stmt as ClassDeclaration, line, character);
        } else if (stmt.type === 'VariableDeclaration') {
            return this.definitionForVariable(stmt as VariableDeclaration, line, character);
        } else if (stmt.type === 'EventDeclaration') {
            return this.definitionForEvent(stmt as EventDeclaration, line, character);
        }

        return null;
    }

    private definitionForProcedure(proc: ProcedureDeclaration, line: number, character: number): LocationInfo | null {
        const procLine = proc.loc!.start.line - 1;
        const procStartCol = proc.loc!.start.column - 1;

        if (line !== procLine) {
            return null;
        }

        // Find keyword length (Sub/Function/Property Get/Let/Set)
        let keyword = proc.isFunction ? 'Function' : 'Sub';
        let keywordEnd = procStartCol + keyword.length;

        if (proc.isProperty) {
            keyword = 'Property';
            keywordEnd = procStartCol + keyword.length;
            const propType = (proc as any).propertyType || 'Get';
            keywordEnd += 1 + propType.length;
        }

        // Name should be after keyword + space
        const nameStart = keywordEnd + 1;
        const nameEnd = nameStart + proc.name.name.length;

        // Check if character is within the procedure name
        if (character >= nameStart && character < nameEnd) {
            return {
                uri: this.uri,
                range: {
                    start: { line: procLine, character: nameStart },
                    end: { line: procLine, character: nameEnd },
                },
            };
        }

        return null;
    }

    private definitionForClass(cls: ClassDeclaration, line: number, character: number): LocationInfo | null {
        // Class declaration doesn't have precise position tracking in loc,
        // so we return a simple range
        return {
            uri: this.uri,
            range: {
                start: { line: 0, character: 0 },
                end: { line: 0, character: cls.name.length },
            },
        };
    }

    private definitionForVariable(decl: VariableDeclaration, line: number, character: number): LocationInfo | null {
        const declLine = (decl.loc?.start.line ?? 1) - 1;

        if (line === declLine) {
            // For "Dim x As Integer", the variable starts after "Dim " keyword
            // loc.start.column is 1-based, convert to 0-based
            const dimStart = (decl.loc?.start.column ?? 1) - 1;
            const dimKeywordLength = 3; // "Dim"
            const varStart = dimStart + dimKeywordLength + 1; // +1 for space

            // Check if character is roughly in the variable name area
            // We return a simple range since we don't track declarator positions precisely
            if (character >= varStart && character < varStart + 20) {
                return {
                    uri: this.uri,
                    range: {
                        start: { line: declLine, character: varStart },
                        end: { line: declLine, character: (decl.loc?.end.column ?? varStart + 10) - 1 },
                    },
                };
            }
        }

        return null;
    }

    private definitionForEvent(evt: EventDeclaration, line: number, character: number): LocationInfo | null {
        const evtLoc = (evt as any).loc;
        if (!evtLoc) return null;

        const evtLine = evtLoc.start.line - 1;

        if (line === evtLine) {
            return {
                uri: this.uri,
                range: {
                    start: { line: evtLine, character: evtLoc.start.column - 1 },
                    end: { line: evtLine, character: evtLoc.end.column - 1 },
                },
            };
        }

        return null;
    }
}
