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
    insertText?: string;
    filterText?: string;
    /** 行内でこの文字位置から cursor までを insertText で置換する */
    replaceStartCharacter?: number;
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

        // Block-closer completions take priority when at line start
        const blockClosers = this.getBlockClosers(source, line, character);
        if (blockClosers.length > 0) return blockClosers;

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

    /**
     * カーソル行が End / Next / Loop / Wend で始まる場合、
     * 上方スキャンして未閉じブロックに対応する閉じキーワードを返す。
     */
    getBlockClosers(source: string, line: number, character: number): CompletionItem[] {
        const lines = source.split('\n');
        if (line >= lines.length) return [];

        const rawCurrentLine = lines[line] ?? '';
        const beforeCursor = rawCurrentLine.substring(0, character);
        const trimmed = beforeCursor.trimStart();
        const indent = beforeCursor.length - trimmed.length;
        const upper = trimmed.toUpperCase();

        // トリガー判定: 行頭（インデント後）が End / Next / Loop / Wend で始まる
        const triggerMatch = upper.match(/^(END|NEXT|LOOP|WEND)(\s|$)/);
        if (!triggerMatch) return [];
        const trigger = triggerMatch[1]; // 'END' | 'NEXT' | 'LOOP' | 'WEND'

        // 上方スキャン: 閉じ済みブロック（後に現れたクローザー）をスタックで管理
        const closersSeen: string[] = [];

        const suggestions: CompletionItem[] = [];

        for (let i = line - 1; i >= 0; i--) {
            const raw = lines[i];
            // 文字列リテラルとコメントを除去（簡易）
            const stripped = raw.replace(/"[^"]*"/g, '""').replace(/'[^"]*$/, '').trim();
            const u = stripped.toUpperCase();

            // ── クローザー（後ろから見て既に閉じているブロック）──
            if (/^END\s+SUB\b/.test(u))      { closersSeen.push('SUB'); continue; }
            if (/^END\s+FUNCTION\b/.test(u)) { closersSeen.push('FUNCTION'); continue; }
            if (/^END\s+PROPERTY\b/.test(u)) { closersSeen.push('PROPERTY'); continue; }
            if (/^END\s+IF\b/.test(u))       { closersSeen.push('IF'); continue; }
            if (/^END\s+WITH\b/.test(u))     { closersSeen.push('WITH'); continue; }
            if (/^END\s+SELECT\b/.test(u))   { closersSeen.push('SELECT'); continue; }
            if (/^END\s+TYPE\b/.test(u))     { closersSeen.push('TYPE'); continue; }
            if (/^END\s+ENUM\b/.test(u))     { closersSeen.push('ENUM'); continue; }
            if (/^NEXT\b/.test(u))           { closersSeen.push('FOR'); continue; }
            if (/^LOOP\b/.test(u))           { closersSeen.push('DO'); continue; }
            if (/^WEND\b/.test(u))           { closersSeen.push('WEND'); continue; }

            // ── オープナー ──
            const makeItem = (label: string): CompletionItem => ({
                label,
                kind: CompletionItemKind.Keyword,
                insertText: label,
                filterText: label,
                replaceStartCharacter: indent,
            });

            const tryMatch = (openerType: string, label: string): boolean => {
                const top = closersSeen[closersSeen.length - 1];
                if (top === openerType) { closersSeen.pop(); return false; } // 閉じ済み
                // 未閉じ: trigger に合うものだけ追加
                if (trigger === 'END' && label.startsWith('End ')) suggestions.push(makeItem(label));
                if (trigger === 'NEXT' && label.startsWith('Next')) suggestions.push(makeItem(label));
                if (trigger === 'LOOP' && label === 'Loop') suggestions.push(makeItem(label));
                if (trigger === 'WEND' && label === 'Wend') suggestions.push(makeItem(label));
                return true; // openerを消費した場合でも走査継続
            };

            // Sub
            if (/^(PUBLIC\s+|PRIVATE\s+|FRIEND\s+|STATIC\s+)*SUB\s+\w/.test(u)) {
                tryMatch('SUB', 'End Sub');
                if (suggestions.length && trigger === 'END') break; // Sub は最外層なのでここで終了
                continue;
            }
            // Function
            if (/^(PUBLIC\s+|PRIVATE\s+|FRIEND\s+|STATIC\s+)*FUNCTION\s+\w/.test(u)) {
                tryMatch('FUNCTION', 'End Function');
                if (suggestions.length && trigger === 'END') break;
                continue;
            }
            // Property
            if (/^(PUBLIC\s+|PRIVATE\s+|FRIEND\s+)*PROPERTY\s+(GET|LET|SET)\s+\w/.test(u)) {
                tryMatch('PROPERTY', 'End Property');
                if (suggestions.length && trigger === 'END') break;
                continue;
            }
            // If (ブロック形式: 行末が Then のみ)
            if (/^IF\b/.test(u) && /\bTHEN\s*$/.test(u)) {
                tryMatch('IF', 'End If');
                continue;
            }
            // For Each / For
            const forEachM = stripped.match(/^for\s+each\s+(\w+)\s+in\b/i);
            const forM     = stripped.match(/^for\s+(\w+)\s*=/i);
            if (forEachM || forM) {
                const v = (forEachM?.[1] ?? forM?.[1]) ?? '';
                tryMatch('FOR', v ? `Next ${v}` : 'Next');
                continue;
            }
            // Do
            if (/^DO(\s|$)/.test(u)) {
                tryMatch('DO', 'Loop');
                continue;
            }
            // While
            if (/^WHILE\s+/.test(u)) {
                tryMatch('WEND', 'Wend');
                continue;
            }
            // With
            if (/^WITH\s+/.test(u)) {
                tryMatch('WITH', 'End With');
                continue;
            }
            // Select Case
            if (/^SELECT\s+CASE\b/.test(u)) {
                tryMatch('SELECT', 'End Select');
                continue;
            }
            // Type
            if (/^(PUBLIC\s+|PRIVATE\s+)*TYPE\s+\w/.test(u)) {
                tryMatch('TYPE', 'End Type');
                if (suggestions.length && trigger === 'END') break;
                continue;
            }
            // Enum
            if (/^(PUBLIC\s+|PRIVATE\s+)*ENUM\s+\w/.test(u)) {
                tryMatch('ENUM', 'End Enum');
                if (suggestions.length && trigger === 'END') break;
                continue;
            }
        }

        return suggestions;
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
