import {
    ClassDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    Statement,
} from '../engine/parser';

export interface CompletionItem {
    label: string;
    kind: CompletionItemKind;
    detail?: string;
    documentation?: string;
}

export enum CompletionItemKind {
    Text = 1,
    Method = 2,
    Function = 3,
    Constructor = 4,
    Field = 5,
    Variable = 6,
    Class = 7,
    Interface = 8,
    Module = 9,
    Property = 10,
    Unit = 11,
    Value = 12,
    Enum = 13,
    Keyword = 14,
    Snippet = 15,
    Color = 16,
    File = 17,
    Reference = 18,
    Folder = 19,
    EnumMember = 20,
    Constant = 21,
    Struct = 22,
    Event = 23,
    Operator = 24,
    TypeParameter = 25,
}

export class CompletionProvider {
    private standardFunctions: CompletionItem[] = [
        // String functions
        { label: 'Len', kind: CompletionItemKind.Function, detail: 'Len(string) As Long' },
        { label: 'LCase', kind: CompletionItemKind.Function, detail: 'LCase(string) As String' },
        { label: 'UCase', kind: CompletionItemKind.Function, detail: 'UCase(string) As String' },
        { label: 'Trim', kind: CompletionItemKind.Function, detail: 'Trim(string) As String' },
        { label: 'Left', kind: CompletionItemKind.Function, detail: 'Left(string, length) As String' },
        { label: 'Right', kind: CompletionItemKind.Function, detail: 'Right(string, length) As String' },
        { label: 'Mid', kind: CompletionItemKind.Function, detail: 'Mid(string, start, [length]) As String' },
        { label: 'Str', kind: CompletionItemKind.Function, detail: 'Str(number) As String' },
        { label: 'Val', kind: CompletionItemKind.Function, detail: 'Val(string) As Double' },
        { label: 'InStr', kind: CompletionItemKind.Function, detail: 'InStr([start,] string1, string2) As Long' },
        { label: 'Replace', kind: CompletionItemKind.Function, detail: 'Replace(expression, find, replace) As String' },
        { label: 'Split', kind: CompletionItemKind.Function, detail: 'Split(expression, [delimiter]) As String()' },
        { label: 'Join', kind: CompletionItemKind.Function, detail: 'Join(sourceArray, [delimiter]) As String' },
        // Numeric functions
        { label: 'Int', kind: CompletionItemKind.Function, detail: 'Int(number) As Long' },
        { label: 'CDbl', kind: CompletionItemKind.Function, detail: 'CDbl(expression) As Double' },
        { label: 'CLng', kind: CompletionItemKind.Function, detail: 'CLng(expression) As Long' },
        { label: 'CInt', kind: CompletionItemKind.Function, detail: 'CInt(expression) As Integer' },
        { label: 'Abs', kind: CompletionItemKind.Function, detail: 'Abs(number) As variant' },
        { label: 'Sqr', kind: CompletionItemKind.Function, detail: 'Sqr(number) As Double' },
        // Type/Variant functions
        { label: 'IsEmpty', kind: CompletionItemKind.Function, detail: 'IsEmpty(expression) As Boolean' },
        { label: 'IsNumeric', kind: CompletionItemKind.Function, detail: 'IsNumeric(expression) As Boolean' },
        { label: 'IsNull', kind: CompletionItemKind.Function, detail: 'IsNull(expression) As Boolean' },
        { label: 'TypeName', kind: CompletionItemKind.Function, detail: 'TypeName(varname) As String' },
        { label: 'VarType', kind: CompletionItemKind.Function, detail: 'VarType(varname) As Integer' },
        // Array functions
        { label: 'UBound', kind: CompletionItemKind.Function, detail: 'UBound(array, [dimension]) As Long' },
        { label: 'LBound', kind: CompletionItemKind.Function, detail: 'LBound(array, [dimension]) As Long' },
        // Other
        { label: 'MsgBox', kind: CompletionItemKind.Function, detail: 'MsgBox(prompt, [buttons], [title]) As Long' },
        { label: 'InputBox', kind: CompletionItemKind.Function, detail: 'InputBox(prompt, [title], [default]) As String' },
        { label: 'CreateObject', kind: CompletionItemKind.Function, detail: 'CreateObject(class) As Object' },
    ];

    getCompletions(statements: Statement[], source: string, line: number, character: number): CompletionItem[] {
        const completions: CompletionItem[] = [];

        // Get prefix at cursor position
        const prefix = this.extractPrefix(source, line, character);

        // Add standard functions
        completions.push(
            ...this.standardFunctions.filter((fn) => this.matchesPrefix(fn.label, prefix))
        );

        // Add local symbols (variables, procedures, classes)
        const symbols = this.extractSymbols(statements);
        completions.push(
            ...symbols.filter((sym) => this.matchesPrefix(sym.label, prefix))
        );

        return completions;
    }

    private extractPrefix(source: string, line: number, character: number): string {
        const lines = source.split('\n');
        if (line >= lines.length) return '';

        const currentLine = lines[line];
        const beforeCursor = currentLine.substring(0, character);

        // Extract identifier: alphanumeric and underscore
        const match = beforeCursor.match(/([a-zA-Z_][a-zA-Z0-9_]*)$/);
        return match ? match[1] : '';
    }

    private matchesPrefix(label: string, prefix: string): boolean {
        if (!prefix) return true;
        return label.toLowerCase().startsWith(prefix.toLowerCase());
    }

    private extractSymbols(statements: Statement[]): CompletionItem[] {
        const symbols: CompletionItem[] = [];

        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                const scope = proc.scope ?? 'public';
                symbols.push({
                    label: proc.name.name,
                    kind: proc.isFunction ? CompletionItemKind.Function : CompletionItemKind.Method,
                    detail: `${scope} ${proc.isFunction ? 'Function' : 'Sub'}`,
                });
            } else if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as ClassDeclaration;
                symbols.push({
                    label: cls.name,
                    kind: CompletionItemKind.Class,
                    detail: 'Class',
                });
                // Add class members
                for (const member of cls.body) {
                    if (member.type === 'VariableDeclaration') {
                        const decl = member as VariableDeclaration;
                        const scope = decl.scope ?? 'private';
                        for (const d of decl.declarations) {
                            symbols.push({
                                label: d.name.name,
                                kind: CompletionItemKind.Property,
                                detail: `${scope} ${d.objectType || 'Variant'} (class member)`,
                            });
                        }
                    } else if (member.type === 'ProcedureDeclaration') {
                        const proc = member as ProcedureDeclaration;
                        const scope = proc.scope ?? 'public';
                        symbols.push({
                            label: proc.name.name,
                            kind: proc.isFunction ? CompletionItemKind.Function : CompletionItemKind.Method,
                            detail: `${scope} ${proc.isFunction ? 'Function' : 'Sub'} (class member)`,
                        });
                    }
                }
            } else if (stmt.type === 'VariableDeclaration') {
                const decl = stmt as VariableDeclaration;
                const scope = decl.scope ?? 'private';
                for (const d of decl.declarations) {
                    symbols.push({
                        label: d.name.name,
                        kind: CompletionItemKind.Variable,
                        detail: `${scope} ${d.objectType || 'Variant'}`,
                    });
                }
            }
        }

        return symbols;
    }
}
