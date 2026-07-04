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
    /** チェーンアクセス解決用: このメンバーの戻り値の型名 (lowercase) */
    returnType?: string;
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

function m(label: string, kind: CompletionItemKind, detail: string, returnType?: string): CompletionItem {
    return returnType ? { label, kind, detail, returnType } : { label, kind, detail };
}
const M = CompletionItemKind.Method;
const P = CompletionItemKind.Property;
const F = CompletionItemKind.Function;

// vba-types.json に出力する際の表示名マッピング (lowercase key → 正式名称)
const TYPE_DISPLAY_NAMES: Record<string, string> = {
    'scripting.dictionary':    'Scripting.Dictionary',
    'scripting.filesystemobject': 'Scripting.FileSystemObject',
    'collection':              'Collection',
    'adodb.recordset':         'ADODB.Recordset',
    'adodb.connection':        'ADODB.Connection',
    'regexp':                  'RegExp',
    'vbscript.regexp':         'VBScript.RegExp',
    'range':                   'Range',
    'worksheet':               'Worksheet',
    'workbook':                'Workbook',
    'sheets':                  'Sheets',
    'application':             'Application',
};

const KIND_NAMES: Record<number, string> = {
    [CompletionItemKind.Method]:    'Method',
    [CompletionItemKind.Function]:  'Function',
    [CompletionItemKind.Property]:  'Property',
    [CompletionItemKind.Variable]:  'Variable',
    [CompletionItemKind.Class]:     'Class',
    [CompletionItemKind.Constant]:  'Constant',
};

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
    // ─── Excel Object Model ──────────────────────────────────────────────────
    ['range', [
        m('Value',          P, 'Value As Variant'),
        m('Value2',         P, 'Value2 As Variant'),
        m('Text',           P, 'Text As String'),
        m('Formula',        P, 'Formula As String'),
        m('FormulaR1C1',    P, 'FormulaR1C1 As String'),
        m('Address',        F, 'Address([RowAbsolute], [ColumnAbsolute], [ReferenceStyle], [External], [RelativeTo]) As String'),
        m('Row',            P, 'Row As Long'),
        m('Column',         P, 'Column As Long'),
        m('Rows',           P, 'Rows As Range', 'range'),
        m('Columns',        P, 'Columns As Range', 'range'),
        m('Cells',          P, 'Cells As Range', 'range'),
        m('Count',          P, 'Count As Long'),
        m('CountLarge',     P, 'CountLarge As Double'),
        m('EntireRow',      P, 'EntireRow As Range', 'range'),
        m('EntireColumn',   P, 'EntireColumn As Range', 'range'),
        m('CurrentRegion',  P, 'CurrentRegion As Range', 'range'),
        m('MergeArea',      P, 'MergeArea As Range', 'range'),
        m('Interior',       P, 'Interior As Interior'),
        m('Font',           P, 'Font As Font'),
        m('Borders',        P, 'Borders As Borders'),
        m('NumberFormat',   P, 'NumberFormat As String'),
        m('HorizontalAlignment', P, 'HorizontalAlignment As XlHAlign'),
        m('VerticalAlignment',   P, 'VerticalAlignment As XlVAlign'),
        m('WrapText',       P, 'WrapText As Boolean'),
        m('MergeCells',     P, 'MergeCells As Boolean'),
        m('Locked',         P, 'Locked As Boolean'),
        m('Hidden',         P, 'Hidden As Boolean'),
        m('Select',         M, 'Select()'),
        m('Activate',       M, 'Activate()'),
        m('Clear',          M, 'Clear()'),
        m('ClearContents',  M, 'ClearContents()'),
        m('ClearFormats',   M, 'ClearFormats()'),
        m('Copy',           M, 'Copy([destination])'),
        m('Cut',            M, 'Cut([destination])'),
        m('Delete',         M, 'Delete([shift])'),
        m('Insert',         M, 'Insert([shift], [copyorigin])'),
        m('Merge',          M, 'Merge([across])'),
        m('UnMerge',        M, 'UnMerge()'),
        m('Sort',           M, 'Sort(...)'),
        m('AutoFit',        M, 'AutoFit()'),
        m('Offset',         F, 'Offset(RowOffset, ColumnOffset) As Range', 'range'),
        m('Resize',         F, 'Resize([RowSize], [ColumnSize]) As Range', 'range'),
        m('End',            F, 'End(direction) As Range', 'range'),
        m('Find',           F, 'Find(what, [after], ...) As Range', 'range'),
        m('FindNext',       F, 'FindNext([after]) As Range', 'range'),
        m('FindPrevious',   F, 'FindPrevious([after]) As Range', 'range'),
        m('SpecialCells',   F, 'SpecialCells(type, [value]) As Range', 'range'),
        m('Cells',          F, 'Cells(row, column) As Range', 'range'),
    ]],
    ['worksheet', [
        m('Name',           P, 'Name As String'),
        m('Index',          P, 'Index As Long'),
        m('CodeName',       P, 'CodeName As String'),
        m('Visible',        P, 'Visible As XlSheetVisibility'),
        m('Cells',          P, 'Cells As Range', 'range'),
        m('Rows',           P, 'Rows As Range', 'range'),
        m('Columns',        P, 'Columns As Range', 'range'),
        m('UsedRange',      P, 'UsedRange As Range', 'range'),
        m('AutoFilter',     P, 'AutoFilter As AutoFilter'),
        m('PageSetup',      P, 'PageSetup As PageSetup'),
        m('Shapes',         P, 'Shapes As Shapes'),
        m('Tab',            P, 'Tab As Tab'),
        m('Parent',         P, 'Parent As Workbook', 'workbook'),
        m('Range',          F, 'Range(cell1, [cell2]) As Range', 'range'),
        m('Activate',       M, 'Activate()'),
        m('Select',         M, 'Select([replace])'),
        m('Delete',         M, 'Delete()'),
        m('Copy',           M, 'Copy([before], [after])'),
        m('Move',           M, 'Move([before], [after])'),
        m('Protect',        M, 'Protect([password], ...)'),
        m('Unprotect',      M, 'Unprotect([password])'),
        m('Calculate',      M, 'Calculate()'),
        m('SetBackgroundPicture', M, 'SetBackgroundPicture(filename)'),
    ]],
    ['workbook', [
        m('Name',           P, 'Name As String'),
        m('FullName',       P, 'FullName As String'),
        m('Path',           P, 'Path As String'),
        m('Saved',          P, 'Saved As Boolean'),
        m('ReadOnly',       P, 'ReadOnly As Boolean'),
        m('Worksheets',     P, 'Worksheets As Sheets', 'sheets'),
        m('Sheets',         P, 'Sheets As Sheets', 'sheets'),
        m('ActiveSheet',    P, 'ActiveSheet As Worksheet', 'worksheet'),
        m('Names',          P, 'Names As Names'),
        m('Connections',    P, 'Connections As Connections'),
        m('Save',           M, 'Save()'),
        m('SaveAs',         M, 'SaveAs([filename], ...)'),
        m('Close',          M, 'Close([savechanges], [filename])'),
        m('Activate',       M, 'Activate()'),
        m('Protect',        M, 'Protect([password], ...)'),
        m('Unprotect',      M, 'Unprotect([password])'),
        m('RefreshAll',     M, 'RefreshAll()'),
        m('PrintOut',       M, 'PrintOut([from], [to], ...)'),
    ]],
    ['sheets', [
        m('Count',  P, 'Count As Long'),
        m('Add',    F, 'Add([before], [after], [count], [type]) As Worksheet', 'worksheet'),
        m('Item',   F, 'Item(index) As Worksheet', 'worksheet'),
        m('Copy',   M, 'Copy([before], [after])'),
        m('Delete', M, 'Delete()'),
        m('Move',   M, 'Move([before], [after])'),
    ]],
    ['application', [
        m('ActiveWorkbook',      P, 'ActiveWorkbook As Workbook', 'workbook'),
        m('ActiveSheet',         P, 'ActiveSheet As Worksheet', 'worksheet'),
        m('ActiveCell',          P, 'ActiveCell As Range', 'range'),
        m('Selection',           P, 'Selection As Object'),
        m('Workbooks',           P, 'Workbooks As Workbooks'),
        m('Worksheets',          P, 'Worksheets As Sheets', 'sheets'),
        m('Sheets',              P, 'Sheets As Sheets', 'sheets'),
        m('Cells',               P, 'Cells As Range', 'range'),
        m('ScreenUpdating',      P, 'ScreenUpdating As Boolean'),
        m('DisplayAlerts',       P, 'DisplayAlerts As Boolean'),
        m('EnableEvents',        P, 'EnableEvents As Boolean'),
        m('Calculation',         P, 'Calculation As XlCalculation'),
        m('StatusBar',           P, 'StatusBar As Variant'),
        m('Caption',             P, 'Caption As String'),
        m('WorksheetFunction',   P, 'WorksheetFunction As WorksheetFunction'),
        m('Range',               F, 'Range(cell1, [cell2]) As Range', 'range'),
        m('Run',                 F, 'Run(macro, ...) As Variant'),
        m('Wait',                M, 'Wait(time)'),
        m('Calculate',           M, 'Calculate()'),
        m('Quit',                M, 'Quit()'),
    ]],
]);

export class CompletionProvider {
    /** workspace の vba-types.json から動的にロードされる外部型定義 */
    private typeStubs = new Map<string, CompletionItem[]>();

    setTypeStubs(stubs: Map<string, CompletionItem[]>): void {
        this.typeStubs = stubs;
    }

    /**
     * カーソル位置が `obj.Member` の Member 上にある場合、そのメンバーの detail 文字列を返す。
     * ホバー情報の提供に使用する。
     */
    getMemberHoverInfo(
        source: string,
        line: number,
        char: number,
        statements: Statement[],
    ): { detail: string; label: string } | null {
        const lines = source.split('\n');
        const currentLine = lines[line] ?? '';

        // カーソル位置の単語を前後から合成する
        const beforeCursor = currentLine.substring(0, char);
        const afterCursor  = currentLine.substring(char);
        const wordBefore   = beforeCursor.match(/([a-zA-Z_][a-zA-Z0-9_]*)$/)?.[1] ?? '';
        const wordAfter    = afterCursor.match(/^([a-zA-Z0-9_]*)/)?.[1] ?? '';
        const memberName   = wordBefore + wordAfter;
        if (!memberName) return null;

        // 直前が `.` であるか確認し、オブジェクト式を取り出す
        const beforeMember = beforeCursor.substring(0, beforeCursor.length - wordBefore.length);
        if (!beforeMember.endsWith('.')) return null;
        const exprBeforeDot = beforeMember.slice(0, -1).trimEnd();

        // 型を解決
        const typeName = this.resolveExprType(exprBeforeDot, statements, source, line)
                      ?? this.findVariableType(statements, exprBeforeDot, line)
                      ?? this.findVariableTypeFromSource(source, exprBeforeDot);
        if (!typeName) return null;

        // メンバーを検索
        const members = this.getMembersForType(typeName, statements);
        const member  = members.find(m => m.label.toLowerCase() === memberName.toLowerCase());
        if (!member?.detail) return null;

        return { label: member.label, detail: member.detail };
    }

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
            let typeName = this.findVariableType(statements, memberAccess.objectName, line)
                        ?? this.findVariableTypeFromSource(source, memberAccess.objectName);

            // チェーンアクセス解決: "objectName" 自体が変数でない場合、カーソル前の式全体を評価
            if (!typeName) {
                const lines = source.split('\n');
                const beforeCursor = (lines[line] ?? '').substring(0, character);
                // memberPrefix と末尾の '.' を取り除いた式を解決する
                const chainExpr = beforeCursor.replace(/\.([a-zA-Z_][a-zA-Z0-9_]*)?$/, '').trimEnd();
                typeName = this.resolveExprType(chainExpr, statements, source, line);
            }

            if (typeName) {
                const members = this.getMembersForType(typeName, statements);
                return members.filter(m => this.matchesPrefix(m.label, memberAccess.memberPrefix));
            }
            return [];
        }

        // With ブロック内の暗黙メンバーアクセス: 行頭が `.prefix` の形
        const withMember = this.detectWithMemberAccess(source, line, character);
        if (withMember !== null) {
            const withObj = this.findEnclosingWithObject(source, line);
            if (withObj) {
                const typeName = this.findVariableType(statements, withObj, line)
                              ?? this.findVariableTypeFromSource(source, withObj);
                if (typeName) {
                    const members = this.getMembersForType(typeName, statements);
                    return members.filter(m => this.matchesPrefix(m.label, withMember));
                }
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

    /** With ブロック内の暗黙アクセス: 行がスペースのみ + `.` + 任意の識別子前置詞 の形なら memberPrefix を返す。 */
    private detectWithMemberAccess(source: string, line: number, char: number): string | null {
        const lines = source.split('\n');
        const beforeCursor = (lines[line] ?? '').substring(0, char);
        // 行頭（インデント可）から `.` のみ、または `.identifier_prefix`
        const match = beforeCursor.match(/^\s*\.([a-zA-Z_][a-zA-Z0-9_]*)?$/);
        if (!match) return null;
        return match[1] ?? '';
    }

    /** source の lineNum 行より上を走査し、最も近い未閉じ With の対象識別子を返す。 */
    private findEnclosingWithObject(source: string, lineNum: number): string | null {
        const lines = source.split('\n');
        let depth = 0;
        for (let i = lineNum - 1; i >= 0; i--) {
            const stripped = lines[i]
                .replace(/"[^"]*"/g, '""')   // 文字列を除去
                .replace(/'.*$/, '')           // コメントを除去
                .trim();
            const upper = stripped.toUpperCase();
            if (/^END\s+WITH\b/.test(upper)) { depth++; continue; }
            if (/^WITH\b/.test(upper)) {
                if (depth > 0) { depth--; continue; }
                // 最も近い未閉じ With に到達
                const m = stripped.match(/^With\s+([a-zA-Z_][a-zA-Z0-9_]*)/i);
                return m ? m[1] : null;
            }
        }
        return null;
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

    /** 型名に対応するメンバー一覧を返す（外部型定義 → 組み込み型 → ユーザー定義クラス の優先順）。 */
    private getMembersForType(typeName: string, statements: Statement[]): CompletionItem[] {
        const lowerType = typeName.replace(/^new\s+/i, '').toLowerCase();
        // "excel.range" → "range" などのモジュール修飾を除いた短縮型名
        const shortType = lowerType.replace(/^[a-z_][a-z0-9_]*\./, '');

        // 外部型定義 (vba-types.json) が最優先 — ユーザーが上書き可能
        const stubMembers = this.typeStubs.get(lowerType) ?? this.typeStubs.get(shortType);
        if (stubMembers && stubMembers.length > 0) return stubMembers;

        // 組み込み型定義
        const builtins = BUILTIN_MEMBERS.get(lowerType) ?? BUILTIN_MEMBERS.get(shortType);
        if (builtins && builtins.length > 0) return builtins;

        // ユーザー定義クラス（複数モジュール全体の statements から探す）
        for (const stmt of statements) {
            if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as ClassDeclaration;
                if (cls.name.toLowerCase() === lowerType || cls.name.toLowerCase() === shortType) {
                    return this.getPublicClassMembers(cls);
                }
            }
        }
        return [];
    }

    /**
     * VBA016 診断用: BUILTIN_MEMBERS と typeStubs の全型名（lowercase）を返す。
     * サーバー側でユーザー定義クラス名を追加して使う。
     */
    getKnownTypeNamesForDiagnostics(): Set<string> {
        const known = new Set<string>();
        for (const [key] of BUILTIN_MEMBERS) {
            known.add(key);
            const short = key.replace(/^[a-z_][a-z0-9_]*\./, '');
            if (short !== key) known.add(short);
        }
        for (const [key] of this.typeStubs) {
            known.add(key);
            const short = key.replace(/^[a-z_][a-z0-9_]*\./, '');
            if (short !== key) known.add(short);
        }
        return known;
    }

    /** 式の型を再帰的に解決する（チェーンアクセス: `ws.Cells` → 'range'）。 */
    private resolveExprType(expr: string, statements: Statement[], source: string, line: number): string | null {
        const trimmed = expr.trim();
        if (!trimmed) return null;

        // 単純識別子
        if (/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(trimmed)) {
            return this.findVariableType(statements, trimmed, line)
                ?? this.findVariableTypeFromSource(source, trimmed);
        }

        // チェーン: greedy で最後の .member を切り出す
        // 例: "ws.Cells(1,1)" → base="ws", member="Cells"
        //     "dict.Items(0).Value" → base="dict.Items(0)", member="Value"
        const chainMatch = trimmed.match(/^(.+)\.([a-zA-Z_][a-zA-Z0-9_]*)(?:\([^)]*\))?$/);
        if (chainMatch) {
            const baseExpr   = chainMatch[1];
            const memberName = chainMatch[2];
            const baseType   = this.resolveExprType(baseExpr, statements, source, line);
            if (baseType) {
                const members = this.getMembersForType(baseType, statements);
                const found   = members.find(mi => mi.label.toLowerCase() === memberName.toLowerCase());
                return found?.returnType ?? null;
            }
        }

        return null;
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

/**
 * BUILTIN_MEMBERS の内容を vba-types.json 形式の JSON 文字列として返す。
 * Quick Fix "Initialize vba-types.json" でワークスペースに書き出す際に使用する。
 */
export function generateDefaultTypeStubsJson(): string {
    const obj: Record<string, Array<{ label: string; kind: string; detail?: string; returnType?: string }>> = {};
    for (const [key, members] of BUILTIN_MEMBERS) {
        const displayName = TYPE_DISPLAY_NAMES[key] ?? key;
        obj[displayName] = members.map(item => {
            const entry: { label: string; kind: string; detail?: string; returnType?: string } = {
                label: item.label,
                kind:  KIND_NAMES[item.kind] ?? 'Method',
            };
            if (item.detail)     entry.detail     = item.detail;
            if (item.returnType) entry.returnType  = item.returnType;
            return entry;
        });
    }
    return JSON.stringify(obj, null, 2);
}
