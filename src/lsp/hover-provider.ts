import {
    ClassDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    EventDeclaration,
    Statement,
} from '../compiler/parser';

export interface Hover {
    contents: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class HoverProvider {
    getHoverInfo(statements: Statement[], line: number, character: number): Hover | null {
        return this.findHoverInStatements(statements, line, character);
    }

    private findHoverInStatements(statements: Statement[], line: number, character: number): Hover | null {
        for (const stmt of statements) {
            const result = this.findHoverInStatement(stmt, line, character);
            if (result) return result;
        }
        return null;
    }

    private findHoverInStatement(stmt: Statement, line: number, character: number): Hover | null {
        if (!stmt.loc) return null;

        // Check if position is within this statement's range
        const stmtLine = stmt.loc.start.line - 1; // Convert to 0-based
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

        // Now check the specific statement type
        if (stmt.type === 'ProcedureDeclaration') {
            return this.hoverForProcedure(stmt as ProcedureDeclaration, line, character);
        } else if (stmt.type === 'ClassDeclaration') {
            return this.hoverForClass(stmt as ClassDeclaration, line, character);
        } else if (stmt.type === 'VariableDeclaration') {
            return this.hoverForVariable(stmt as VariableDeclaration, line, character);
        } else if (stmt.type === 'EventDeclaration') {
            return this.hoverForEvent(stmt as EventDeclaration, line, character);
        }

        return null;
    }

    private hoverForProcedure(proc: ProcedureDeclaration, line: number, character: number): Hover | null {
        // The procedure starts at loc.start (Sub/Function/Property keyword)
        // The name comes after: "Sub Foo" means 'Foo' starts 4 chars into the line
        // loc.start.column is 1-based, so we need to find where the name is
        // For "Sub Foo", name is at column 5 (1-based) = character 4 (0-based)

        const procLine = proc.loc!.start.line - 1; // Convert to 0-based
        const procStartCol = proc.loc!.start.column - 1; // Convert to 0-based

        // Check if we're on the same line as the procedure declaration
        if (line !== procLine) {
            return null;
        }

        // Find keyword length (Sub/Function/Property Get/Let/Set)
        let keyword = proc.isFunction ? 'Function' : 'Sub';
        let keywordEnd = procStartCol + keyword.length;

        if (proc.isProperty) {
            keyword = 'Property';
            keywordEnd = procStartCol + keyword.length;
            // For Property, add " Get"/"Let"/"Set" to the keyword
            const propType = (proc as any).propertyType || 'Get';
            keywordEnd += 1 + propType.length; // +1 for space
        }

        // Name should be after keyword + space
        const nameStart = keywordEnd + 1;
        const nameEnd = nameStart + proc.name.name.length;

        // Check if character is within the procedure name
        if (character >= nameStart && character < nameEnd) {
            const params = proc.parameters.map(p => `${p.name.name} As ${p.typeName || 'Variant'}`).join(', ');
            const sig = params ? `${keyword} ${proc.name.name}(${params})` : `${keyword} ${proc.name.name}()`;
            const returnType = proc.isFunction || proc.isProperty ? ` As ${proc.name.name || 'Variant'}` : '';

            return {
                contents: sig + returnType,
                range: {
                    start: { line: procLine, character: nameStart },
                    end: { line: procLine, character: nameEnd },
                },
            };
        }

        return null;
    }

    private hoverForClass(cls: ClassDeclaration, line: number, character: number): Hover | null {
        const nameStart = cls.name.toLowerCase() === 'class' ? 6 : 0; // Rough estimate
        const nameEnd = nameStart + cls.name.length;

        // Check if position is near the class name (line typically doesn't have column info for classes in AST)
        // For simplicity, check if line is within class declaration line
        if (line === 0 || cls.body.length === 0) {
            return {
                contents: `Class ${cls.name}`,
                range: {
                    start: { line: 0, character: nameStart },
                    end: { line: 0, character: nameEnd },
                },
            };
        }

        // Check members
        const memberHover = this.findHoverInStatements(cls.body, line, character);
        if (memberHover) {
            return memberHover;
        }

        return null;
    }

    private hoverForVariable(decl: VariableDeclaration, line: number, character: number): Hover | null {
        for (const d of decl.declarations) {
            const nameStart = d.name.name.length ? 0 : 0; // Variable name position not precisely tracked
            const typeName = d.typeName || 'Variant';
            const info = `Dim ${d.name.name} As ${typeName}`;

            // Since we don't have precise column tracking for variable names in declarators,
            // just return if line matches
            if (line === (decl.loc?.start.line ?? 1) - 1) {
                return {
                    contents: info,
                    range: {
                        start: { line: decl.loc?.start.line ?? 1 - 1, character: 0 },
                        end: { line: decl.loc?.end.line ?? 1 - 1, character: decl.loc?.end.column ?? 10 },
                    },
                };
            }
        }

        return null;
    }

    private hoverForEvent(evt: EventDeclaration, line: number, character: number): Hover | null {
        const name = (evt as any).name?.name || 'Event';
        const params = ((evt as any).parameters || [])
            .map((p: any) => `${p.name.name} As ${p.typeName || 'Variant'}`)
            .join(', ');
        const sig = params ? `Event ${name}(${params})` : `Event ${name}()`;

        return {
            contents: sig,
            range: {
                start: { line: (evt as any).loc?.start.line ?? 0 - 1, character: 0 },
                end: { line: (evt as any).loc?.end.line ?? 0 - 1, character: 10 },
            },
        };
    }
}
