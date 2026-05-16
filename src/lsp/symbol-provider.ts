import {
    ClassDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    Statement,
    SourceLocation
} from '../compiler/parser';

export interface DocumentSymbol {
    name: string;
    kind: SymbolKind;
    location: Location;
    containerName?: string;
    children?: DocumentSymbol[];
}

export interface Location {
    uri?: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export enum SymbolKind {
    File = 1,
    Module = 2,
    Namespace = 3,
    Package = 4,
    Class = 5,
    Method = 6,
    Property = 7,
    Field = 8,
    Constructor = 9,
    Enum = 10,
    Interface = 11,
    Function = 12,
    Variable = 13,
    Constant = 14,
    String = 15,
    Number = 16,
    Boolean = 17,
    Array = 18,
    Object = 19,
    Key = 20,
    Null = 21,
    EnumMember = 22,
    Struct = 23,
    Event = 24,
    Operator = 25,
    TypeParameter = 26,
}

export class SymbolProvider {
    private uri: string = '';

    setDocumentUri(uri: string): void {
        this.uri = uri;
    }

    extractSymbols(statements: Statement[]): DocumentSymbol[] {
        return statements.flatMap((stmt) => this.symbolFromStatement(stmt));
    }

    private symbolFromStatement(stmt: Statement): DocumentSymbol[] {
        if (stmt.type === 'ProcedureDeclaration') {
            return [this.symbolFromProcedure(stmt as ProcedureDeclaration)];
        } else if (stmt.type === 'ClassDeclaration') {
            return [this.symbolFromClass(stmt as ClassDeclaration)];
        } else if (stmt.type === 'VariableDeclaration') {
            return this.symbolsFromVariable(stmt as VariableDeclaration);
        } else if (stmt.type === 'EventDeclaration') {
            const evt = stmt as any;
            return [
                {
                    name: evt.name.name,
                    kind: SymbolKind.Event,
                    location: this.locationFromNode(stmt),
                }
            ];
        }
        return [];
    }

    private symbolFromProcedure(proc: ProcedureDeclaration): DocumentSymbol {
        const kind = proc.isProperty
            ? SymbolKind.Property
            : proc.isFunction
            ? SymbolKind.Function
            : SymbolKind.Method;

        return {
            name: proc.name.name,
            kind,
            location: this.locationFromNode(proc),
        };
    }

    private symbolFromClass(cls: ClassDeclaration): DocumentSymbol {
        const children = cls.body.flatMap((stmt) => {
            const symbols = this.symbolFromStatement(stmt);
            return symbols.map((sym) => ({
                ...sym,
                containerName: cls.name,
            }));
        });

        return {
            name: cls.name,
            kind: SymbolKind.Class,
            location: {
                uri: this.uri,
                range: {
                    start: { line: 0, character: 0 },
                    end: { line: 0, character: 0 },
                },
            },
            children,
        };
    }

    private symbolsFromVariable(decl: VariableDeclaration): DocumentSymbol[] {
        return decl.declarations.map((d) => ({
            name: d.name.name,
            kind: SymbolKind.Variable,
            location: this.locationFromNode(decl),
        }));
    }

    private locationFromNode(node: Statement): Location {
        const range = node.loc
            ? {
                start: { line: node.loc.start.line - 1, character: node.loc.start.column - 1 },
                end: { line: node.loc.end.line - 1, character: node.loc.end.column - 1 },
            }
            : {
                start: { line: 0, character: 0 },
                end: { line: 0, character: 0 },
            };

        return {
            uri: this.uri,
            range,
        };
    }
}
