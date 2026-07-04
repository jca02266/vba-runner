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
    sortText?: string;
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

// ─── 組み込み COM オブジェクトのメンバー定義 ───────────────────────────────────

function m(label: string, kind: CompletionItemKind, detail: string): CompletionItem {
    return { label, kind, detail };
}
const M = CompletionItemKind.Method;
const P = CompletionItemKind.Property;
const F = CompletionItemKind.Function;

const BUILTIN_MEMBERS: ReadonlyMap<string, CompletionItem[]> = new Map([
    ['scripting.dictionary', [
        m('Add',         M, 'Add(key, item)'),
        m('Exists',      F, 'Exists(key) As Boolean'),
        m('Item',        P, 'Item(key) As Variant'),
        m('Items',       F, 'Items() As Variant()'),
        m('Keys',        F, 'Keys() As Variant()'),
        m('Count',       P, 'Count As Long'),
        m('Remove',      M, 'Remove(key)'),
        m('RemoveAll',   M, 'RemoveAll()'),
        m('CompareMode', P, 'CompareMode As Integer'),
    ]],
    ['scripting.filesystemobject', [
        m('FileExists',      F, 'FileExists(filespec) As Boolean'),
        m('FolderExists',    F, 'FolderExists(folderspec) As Boolean'),
        m('DriveExists',     F, 'DriveExists(drivespec) As Boolean'),
        m('GetFile',         F, 'GetFile(filespec) As File'),
        m('GetFolder',       F, 'GetFolder(folderspec) As Folder'),
        m('GetDrive',        F, 'GetDrive(drivespec) As Drive'),
        m('GetFileName',     F, 'GetFileName(path) As String'),
        m('GetBaseName',     F, 'GetBaseName(path) As String'),
        m('GetExtensionName',F, 'GetExtensionName(path) As String'),
        m('GetParentFolderName', F, 'GetParentFolderName(path) As String'),
        m('BuildPath',       F, 'BuildPath(path, name) As String'),
        m('CreateTextFile',  F, 'CreateTextFile(filename, [overwrite], [unicode]) As TextStream'),
        m('OpenTextFile',    F, 'OpenTextFile(filename, [iomode], [create], [format]) As TextStream'),
        m('CopyFile',        M, 'CopyFile(source, destination, [overwrite])'),
        m('MoveFile',        M, 'MoveFile(source, destination)'),
        m('DeleteFile',      M, 'DeleteFile(filespec, [force])'),
        m('CreateFolder',    M, 'CreateFolder(folderspec)'),
        m('CopyFolder',      M, 'CopyFolder(source, destination, [overwrite])'),
        m('MoveFolder',      M, 'MoveFolder(source, destination)'),
        m('DeleteFolder',    M, 'DeleteFolder(folderspec, [force])'),
        m('GetSpecialFolder',F, 'GetSpecialFolder(folderspec) As Folder'),
        m('GetTempName',     F, 'GetTempName() As String'),
        m('Drives',          P, 'Drives As Drives'),
    ]],
    ['collection', [
        m('Add',    M, 'Add(item, [key], [before], [after])'),
        m('Remove', M, 'Remove(index)'),
        m('Item',   F, 'Item(index) As Variant'),
        m('Count',  P, 'Count As Long'),
    ]],
    ['adodb.recordset', [
        m('Open',       M, 'Open([source], [activeconnection], [cursortype], [locktype], [options])'),
        m('Close',      M, 'Close()'),
        m('MoveNext',   M, 'MoveNext()'),
        m('MovePrevious',M,'MovePrevious()'),
        m('MoveFirst',  M, 'MoveFirst()'),
        m('MoveLast',   M, 'MoveLast()'),
        m('AddNew',     M, 'AddNew([fieldlist], [values])'),
        m('Update',     M, 'Update([fieldlist], [values])'),
        m('Delete',     M, 'Delete([affectrecords])'),
        m('Find',       M, 'Find(criteria, [skiprows], [searchdirection], [start])'),
        m('EOF',        P, 'EOF As Boolean'),
        m('BOF',        P, 'BOF As Boolean'),
        m('Fields',     P, 'Fields As Fields'),
        m('RecordCount',P, 'RecordCount As Long'),
        m('Filter',     P, 'Filter As Variant'),
        m('Sort',       P, 'Sort As String'),
        m('GetRows',    F, 'GetRows([rows], [start], [fields]) As Variant()'),
        m('GetString',  F, 'GetString([stringformat], [rows], [columndelimeter], [rowdelimeter], [nullexpr]) As String'),
    ]],
    ['adodb.connection', [
        m('Open',           M, 'Open([connectionstring], [userid], [password], [options])'),
        m('Close',          M, 'Close()'),
        m('Execute',        F, 'Execute(commandtext, [recordsaffected], [options]) As Recordset'),
        m('BeginTrans',     F, 'BeginTrans() As Long'),
        m('CommitTrans',    M, 'CommitTrans()'),
        m('RollbackTrans',  M, 'RollbackTrans()'),
        m('ConnectionString',P,'ConnectionString As String'),
        m('State',          P, 'State As Long'),
        m('Errors',         P, 'Errors As Errors'),
    ]],
    ['regexp', [
        m('Execute',    F, 'Execute(string) As MatchCollection'),
        m('Test',       F, 'Test(string) As Boolean'),
        m('Replace',    F, 'Replace(string, replacement) As String'),
        m('Pattern',    P, 'Pattern As String'),
        m('Global',     P, 'Global As Boolean'),
        m('IgnoreCase', P, 'IgnoreCase As Boolean'),
        m('MultiLine',  P, 'MultiLine As Boolean'),
    ]],
    ['vbscript.regexp', [
        m('Execute',    F, 'Execute(string) As MatchCollection'),
        m('Test',       F, 'Test(string) As Boolean'),
        m('Replace',    F, 'Replace(string, replacement) As String'),
        m('Pattern',    P, 'Pattern As String'),
        m('Global',     P, 'Global As Boolean'),
        m('IgnoreCase', P, 'IgnoreCase As Boolean'),
        m('MultiLine',  P, 'MultiLine As Boolean'),
    ]],
]);

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
        // Block-closer completions take priority when at line start
        const blockClosers = this.getBlockClosers(source, line, character);
        if (blockClosers.length > 0) return blockClosers;

        // Member access: "obj." or "obj.prefix"
        const memberAccess = this.detectMemberAccess(source, line, character);
        if (memberAccess) {
            // AST walk first; fallback to source text scan when AST is incomplete (e.g. mid-type error recovery)
            const typeName = this.findVariableType(statements, memberAccess.objectName, line)
                          ?? this.findVariableTypeFromSource(source, memberAccess.objectName);
            if (typeName) {
                const members = this.getMembersForType(typeName, statements);
                return members.filter(m => this.matchesPrefix(m.label, memberAccess.memberPrefix));
            }
            return [];
        }

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

    /** カーソル直前が `ident.` または `ident.prefix` の形なら検出する。 */
    private detectMemberAccess(source: string, line: number, char: number): { objectName: string; memberPrefix: string } | null {
        const lines = source.split('\n');
        const beforeCursor = (lines[line] ?? '').substring(0, char);
        const match = beforeCursor.match(/([a-zA-Z_][a-zA-Z0-9_]*)\.([a-zA-Z_][a-zA-Z0-9_]*)?$/);
        if (!match) return null;
        return { objectName: match[1], memberPrefix: match[2] ?? '' };
    }

    /** AST を走査して varName の型を返す（手続き内ローカル優先、次にモジュールレベル）。 */
    private findVariableType(statements: Statement[], varName: string, lineNum: number): string | null {
        const lowerVar = varName.toLowerCase();

        // 現在のカーソル行を含むプロシージャを探す
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                const startLine = (proc.loc?.start.line ?? 1) - 1;
                const endLine   = (proc.loc?.end.line   ?? 0) - 1;
                if (lineNum < startLine || lineNum > endLine) continue;

                // パラメーター
                for (const param of proc.parameters) {
                    if (param.name.toLowerCase() === lowerVar && param.paramType) {
                        return param.paramType.toLowerCase();
                    }
                }
                // ローカル Dim
                for (const s of proc.body) {
                    if (s.type === 'VariableDeclaration') {
                        for (const d of (s as VariableDeclaration).declarations) {
                            if (d.name.name.toLowerCase() === lowerVar && d.objectType) {
                                return d.objectType.toLowerCase();
                            }
                        }
                    }
                }
            }
        }

        // モジュールレベル Dim / Public / Private
        for (const stmt of statements) {
            if (stmt.type === 'VariableDeclaration') {
                for (const d of (stmt as VariableDeclaration).declarations) {
                    if (d.name.name.toLowerCase() === lowerVar && d.objectType) {
                        return d.objectType.toLowerCase();
                    }
                }
            }
        }

        return null;
    }

    /** AST が不完全なときのフォールバック: ソーステキストから Dim varName As Type を正規表現で探す。 */
    private findVariableTypeFromSource(source: string, varName: string): string | null {
        const escaped = varName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const dimRe = new RegExp(
            `(?:Dim|Private|Public|Friend|Static)\\s+${escaped}\\s+As(?:\\s+New)?\\s+([a-zA-Z_][a-zA-Z0-9_.]*)`,
            'i',
        );
        const m = source.match(dimRe);
        if (m) return m[1].toLowerCase();
        // Set x = New TypeName または Set x = CreateObject("Type.Name")
        const setRe = new RegExp(`Set\\s+${escaped}\\s*=\\s*(?:New\\s+)?([a-zA-Z_][a-zA-Z0-9_.]*)`, 'i');
        const m2 = source.match(setRe);
        return m2 ? m2[1].toLowerCase() : null;
    }

    /** 型名に対応するメンバー一覧を返す（組み込み型 + ユーザー定義クラス）。 */
    private getMembersForType(typeName: string, statements: Statement[]): CompletionItem[] {
        const lowerType = typeName.replace(/^new\s+/i, '').toLowerCase();
        const builtins = BUILTIN_MEMBERS.get(lowerType);
        if (builtins) return builtins;

        // ユーザー定義クラス
        for (const stmt of statements) {
            if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as ClassDeclaration;
                if (cls.name.toLowerCase() === lowerType) {
                    return this.getPublicClassMembers(cls);
                }
            }
        }
        return [];
    }

    /** クラスの Public メンバーを CompletionItem に変換する。 */
    private getPublicClassMembers(cls: ClassDeclaration): CompletionItem[] {
        const items: CompletionItem[] = [];
        for (const member of cls.body) {
            if (member.type === 'VariableDeclaration') {
                const decl = member as VariableDeclaration;
                if ((decl.scope ?? 'private') === 'private') continue;
                for (const d of decl.declarations) {
                    items.push({
                        label: d.name.name,
                        kind: CompletionItemKind.Property,
                        detail: `${d.objectType ?? 'Variant'}`,
                    });
                }
            } else if (member.type === 'ProcedureDeclaration') {
                const proc = member as ProcedureDeclaration;
                if ((proc.scope ?? 'public') === 'private') continue;
                items.push({
                    label: proc.name.name,
                    kind: proc.isProperty ? CompletionItemKind.Property : proc.isFunction ? CompletionItemKind.Function : CompletionItemKind.Method,
                    detail: proc.isProperty ? 'Property' : proc.isFunction ? 'Function' : 'Sub',
                });
            }
        }
        return items;
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
            // sortText は発見順（内側ほど小さい）で付与し VS Code の並び替えを防ぐ
            const makeItem = (label: string): CompletionItem => ({
                label,
                kind: CompletionItemKind.Keyword,
                insertText: label,
                filterText: label,
                sortText: String(suggestions.length).padStart(4, '0'),
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
