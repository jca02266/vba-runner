/**
 * Word Application モック
 *
 * `injectWordStub(evaluator)` と組み合わせて使うと、Word 依存コードを
 * エンジン上で実行できる。
 *
 * 対応する VBA パターン（代表例）:
 *   ActiveDocument.Content.Text = "Hello"
 *   Selection.TypeText "World"
 *   ActiveDocument.Bookmarks("bmkName").Range.Text = "value"
 *   For Each p In ActiveDocument.Paragraphs: Debug.Print p.Range.Text: Next p
 *   ActiveDocument.Tables(1).Cell(1, 1).Range.Text = "Header"
 *   With Selection.Find: .Text = "old": .Replacement.Text = "new": .Execute: End With
 *   Documents.Open "C:\path\file.docx"
 */

import { VbaType, VbaExternalObject } from '../vba-types';

// ========== フォント・書式スタブ ==========

class MockWordFont {
    get Bold() { return false; }
    set Bold(_v: any) {}
    get Italic() { return false; }
    set Italic(_v: any) {}
    get Underline() { return 0; }
    set Underline(_v: any) {}
    get Color() { return 0; }
    set Color(_v: any) {}
    get Size() { return 12; }
    set Size(_v: any) {}
    get Name() { return 'Calibri'; }
    set Name(_v: any) {}
}

class MockParagraphFormat {
    get Alignment() { return 0; } // wdAlignParagraphLeft
    set Alignment(_v: any) {}
    get LeftIndent() { return 0; }
    set LeftIndent(_v: any) {}
    get RightIndent() { return 0; }
    set RightIndent(_v: any) {}
    get SpaceBefore() { return 0; }
    set SpaceBefore(_v: any) {}
    get SpaceAfter() { return 0; }
    set SpaceAfter(_v: any) {}
    get LineSpacing() { return 0; }
    set LineSpacing(_v: any) {}
}

// ========== Find / Replacement ==========

class MockReplacement {
    Text: string = '';
    ClearFormatting() {}
}

export class MockFind {
    Text: string = '';
    get Replacement(): MockReplacement { return this._replacement; }
    private _replacement = new MockReplacement();
    Forward: boolean = true;
    MatchCase: boolean = false;
    MatchWholeWord: boolean = false;
    MatchWildcards: boolean = false;
    Wrap: number = 1; // wdFindContinue
    Found: boolean = false;

    /** Execute(FindText?, MatchCase?, ..., Replace?, ...) */
    Execute(..._args: any[]): boolean {
        this.Found = false;
        return false; // スタブは常に「見つからない」
    }

    ClearFormatting() {}
}

// ========== Word Range ==========

export class MockWordRange implements VbaType {
    readonly __vbaTypeName__ = 'Range';

    private _text: string;
    private _find: MockFind = new MockFind();

    constructor(text: string = '') {
        this._text = text;
    }

    get Text(): string { return this._text; }
    set Text(v: any) { this._text = String(v ?? ''); }

    get Start(): number { return 0; }
    get End(): number { return this._text.length; }

    /** r.InsertAfter "text" */
    InsertAfter(text: string) { this._text += String(text ?? ''); }
    /** r.InsertBefore "text" */
    InsertBefore(text: string) { this._text = String(text ?? '') + this._text; }

    get Find(): MockFind { return this._find; }

    get Bold() { return false; }
    set Bold(_v: any) {}
    get Italic() { return false; }
    set Italic(_v: any) {}
    get Underline() { return 0; }
    set Underline(_v: any) {}
    get Style() { return ''; }
    set Style(_v: any) {}

    get Font(): MockWordFont { return new MockWordFont(); }
    get ParagraphFormat(): MockParagraphFormat { return new MockParagraphFormat(); }

    Select() {}
    Copy() {}
    Cut() {}
    Paste() {}
    Delete() {}
    SetRange() {}
    Collapse() {}

    get Characters(): MockWordRangeCollection { return new MockWordRangeCollection([]); }
    get Words(): MockWordRangeCollection { return new MockWordRangeCollection([]); }
    get Sentences(): MockWordRangeCollection { return new MockWordRangeCollection([]); }
    get Paragraphs(): MockParagraphsCollection { return new MockParagraphsCollection([]); }
}

// ========== Range コレクション（Characters / Words / Sentences）==========

class MockWordRangeCollection {
    items: MockWordRange[];
    constructor(items: MockWordRange[]) { this.items = items; }
    get Count(): number { return this.items.length; }
    Item(n: number): MockWordRange { return this.items[n - 1] ?? new MockWordRange(); }
}

// ========== Paragraph ==========

export class MockParagraph implements VbaType {
    readonly __vbaTypeName__ = 'Paragraph';
    private _range: MockWordRange;

    constructor(text: string = '') {
        this._range = new MockWordRange(text);
    }

    get Range(): MockWordRange { return this._range; }

    get Style() { return 'Normal'; }
    set Style(_v: any) {}
    get Alignment() { return 0; }
    set Alignment(_v: any) {}
    get LeftIndent() { return 0; }
    set LeftIndent(_v: any) {}
}

// ========== Paragraphs コレクション ==========

export class MockParagraphsCollection {
    items: MockParagraph[];

    constructor(paragraphs: MockParagraph[]) {
        this.items = paragraphs;
    }

    get Count(): number { return this.items.length; }
    Item(n: number): MockParagraph { return this.items[n - 1] ?? new MockParagraph(); }
}

// ========== Table ==========

export class MockTableCell implements VbaType {
    readonly __vbaTypeName__ = 'Cell';
    private _range: MockWordRange;

    constructor(text: string = '') {
        this._range = new MockWordRange(text);
    }

    get Range(): MockWordRange { return this._range; }
    get Width() { return 72; }
    set Width(_v: any) {}
    get VerticalAlignment() { return 0; }
    set VerticalAlignment(_v: any) {}

    Merge() {}
    Split() {}
}

export class MockTable implements VbaType {
    readonly __vbaTypeName__ = 'Table';
    private _cells: MockTableCell[][];

    constructor(rows: number = 3, cols: number = 3) {
        this._cells = Array.from({ length: rows }, () =>
            Array.from({ length: cols }, () => new MockTableCell())
        );
    }

    Cell(row: number, col: number): MockTableCell {
        const r = this._cells[row - 1];
        return r ? (r[col - 1] ?? new MockTableCell()) : new MockTableCell();
    }

    get Rows(): { Count: number; Item: (n: number) => any } {
        const n = this._cells.length;
        return { Count: n, Item: (_i: number) => ({ Cells: { Count: this._cells[0]?.length ?? 0 } }) };
    }

    get Columns(): { Count: number; Item: (n: number) => any } {
        const n = this._cells[0]?.length ?? 0;
        return { Count: n, Item: (_i: number) => ({}) };
    }

    get Range(): MockWordRange { return new MockWordRange(); }

    Select() {}
    Delete() {}
    AutoFit() {}
}

// ========== Tables コレクション ==========

class MockTablesCollection {
    items: MockTable[];

    constructor(tables: MockTable[]) {
        this.items = tables;
    }

    get Count(): number { return this.items.length; }
    Item(n: number): MockTable { return this.items[n - 1] ?? new MockTable(); }
}

// ========== Bookmark ==========

export class MockBookmark implements VbaType {
    readonly __vbaTypeName__ = 'Bookmark';
    private _range: MockWordRange;
    readonly Name: string;

    constructor(name: string, text: string = '') {
        this.Name = name;
        this._range = new MockWordRange(text);
    }

    get Range(): MockWordRange { return this._range; }
    get Exists(): boolean { return true; }
}

// ========== Bookmarks コレクション ==========

class MockBookmarksCollection {
    private _store: Map<string, MockBookmark> = new Map();

    /** テスト用: ブックマークを事前登録する */
    add(name: string, text: string = ''): MockBookmark {
        const bm = new MockBookmark(name, text);
        this._store.set(name.toLowerCase(), bm);
        return bm;
    }

    Item(name: string): MockBookmark {
        const key = typeof name === 'string' ? name.toLowerCase() : String(name);
        if (!this._store.has(key)) this._store.set(key, new MockBookmark(String(name)));
        return this._store.get(key)!;
    }

    Exists(name: string): boolean {
        return this._store.has(name.toLowerCase());
    }

    get Count(): number { return this._store.size; }
    get items(): MockBookmark[] { return [...this._store.values()]; }
}

// ========== Selection ==========

export class MockSelection implements VbaType {
    readonly __vbaTypeName__ = 'Selection';
    private _text: string = '';
    private _range: MockWordRange;
    private _find: MockFind = new MockFind();

    constructor() {
        this._range = new MockWordRange('');
    }

    get Text(): string { return this._text; }
    set Text(v: any) { this._text = String(v ?? ''); this._range.Text = this._text; }

    get Range(): MockWordRange { return this._range; }
    get Find(): MockFind { return this._find; }

    /** VBA: Selection.TypeText "Hello" */
    TypeText(text: string) {
        this._text += String(text ?? '');
        this._range.Text = this._text;
    }

    TypeParagraph() { this._text += '\n'; this._range.Text = this._text; }
    TypeBackspace() { this._text = this._text.slice(0, -1); }

    MoveRight() {}
    MoveLeft() {}
    MoveUp() {}
    MoveDown() {}
    MoveStart() {}
    MoveEnd() {}
    HomeKey() {}
    EndKey() {}

    get Bold() { return false; }
    set Bold(_v: any) {}
    get Italic() { return false; }
    set Italic(_v: any) {}
    get Underline() { return 0; }
    set Underline(_v: any) {}
    get Style() { return 'Normal'; }
    set Style(_v: any) {}

    get Font(): MockWordFont { return new MockWordFont(); }
    get ParagraphFormat(): MockParagraphFormat { return new MockParagraphFormat(); }

    Select() {}
    Copy() {}
    Cut() {}
    Paste() {}
    Delete() {}
    ClearFormatting() {}
    WholeStory() {}
}

// ========== Document ==========

export class MockDocument implements VbaType {
    readonly __vbaTypeName__ = 'Document';

    private _name: string;
    private _path: string;
    private _content: MockWordRange;
    private _paragraphs: MockParagraphsCollection;
    private _tables: MockTablesCollection;
    private _bookmarks: MockBookmarksCollection;

    constructor(name: string = 'MockDocument.docx', path: string = '') {
        this._name = name;
        this._path = path;
        this._content = new MockWordRange('');
        this._paragraphs = new MockParagraphsCollection([]);
        this._tables = new MockTablesCollection([]);
        this._bookmarks = new MockBookmarksCollection();
    }

    get Name(): string { return this._name; }
    get FullName(): string { return this._path ? `${this._path}\\${this._name}` : this._name; }
    get Path(): string { return this._path; }
    get Saved(): boolean { return true; }
    get ReadOnly(): boolean { return false; }

    /** doc.Content → 文書全体の Range */
    get Content(): MockWordRange { return this._content; }

    /** doc.Range(start?, end?) */
    Range(_start?: number, _end?: number): MockWordRange {
        return new MockWordRange(this._content.Text);
    }

    /** doc.Paragraphs / doc.Paragraphs(n) */
    Paragraphs(...args: any[]): MockParagraphsCollection | MockParagraph {
        if (args.length > 0) return this._paragraphs.Item(Number(args[0]));
        return this._paragraphs;
    }

    /** doc.Tables / doc.Tables(n) */
    Tables(...args: any[]): MockTablesCollection | MockTable {
        if (args.length > 0) return this._tables.Item(Number(args[0]));
        return this._tables;
    }

    /** doc.Bookmarks("name") / doc.Bookmarks.Item("name") */
    Bookmarks(...args: any[]): MockBookmarksCollection | MockBookmark {
        if (args.length > 0) return this._bookmarks.Item(String(args[0]));
        return this._bookmarks;
    }

    Save() {}
    SaveAs(_filename?: string) {}
    Close(_saveChanges?: any) {}
    Protect() {}
    Unprotect() {}
    Activate() {}
    PrintOut() {}

    // ==============================
    // テスト用ユーティリティ
    // ==============================

    /** 本文テキストを設定する（段落は '\n' で区切る） */
    setContent(text: string): void {
        this._content.Text = text;
        this._paragraphs = new MockParagraphsCollection(
            text.split('\n').map(t => new MockParagraph(t))
        );
    }

    /** テーブルを追加する */
    addTable(rows: number = 3, cols: number = 3): MockTable {
        const tbl = new MockTable(rows, cols);
        this._tables.items.push(tbl);
        return tbl;
    }

    /** ブックマークへのアクセス（事前登録用） */
    get bookmarks(): MockBookmarksCollection { return this._bookmarks; }
}

// ========== Documents コレクション ==========

class MockDocumentsCollection {
    private _docs: MockDocument[] = [];

    get Count(): number { return this._docs.length; }
    get items(): MockDocument[] { return this._docs; }

    Item(n: number): MockDocument { return this._docs[n - 1] ?? new MockDocument(); }

    Open(_filename?: string): MockDocument {
        const name = _filename ? String(_filename).split(/[\\/]/).pop()! : 'MockDocument.docx';
        const doc = new MockDocument(name, _filename ?? '');
        this._docs.push(doc);
        return doc;
    }

    Add(): MockDocument {
        const doc = new MockDocument(`Document${this._docs.length + 1}.docx`);
        this._docs.push(doc);
        return doc;
    }
}

// ========== Word Application ==========

export class MockWordApplication implements VbaExternalObject {
    readonly __className__ = 'Word.Application';

    private _activeDoc: MockDocument;
    private _documents: MockDocumentsCollection;
    private _selection: MockSelection;

    constructor(docName: string = 'Document1.docx') {
        this._activeDoc = new MockDocument(docName);
        this._documents = new MockDocumentsCollection();
        this._documents['_docs'].push(this._activeDoc);
        this._selection = new MockSelection();
    }

    // ==============================
    // Tier 6 グローバルメンバー
    // ==============================

    /** VBA: ActiveDocument.Content.Text */
    get ActiveDocument(): MockDocument { return this._activeDoc; }

    /** VBA: Selection.TypeText "hello" */
    get Selection(): MockSelection { return this._selection; }

    /** VBA: Documents.Open "C:\path\file.docx" / Documents.Count */
    Documents(...args: any[]): MockDocumentsCollection | MockDocument {
        if (args.length > 0) return this._documents.Item(Number(args[0]));
        return this._documents;
    }

    // ==============================
    // Application レベルの設定（ノーオプ）
    // ==============================

    get Name(): string { return 'Microsoft Word'; }
    get Version(): string { return '16.0'; }
    get Application(): this { return this; }

    get ScreenUpdating() { return true; }
    set ScreenUpdating(_v: any) {}
    get DisplayAlerts() { return 0; }
    set DisplayAlerts(_v: any) {}
    get Visible() { return true; }
    set Visible(_v: any) {}

    Quit() {}

    // ==============================
    // テスト用ユーティリティ
    // ==============================

    /** ActiveDocument へのショートカット（事前データ設定用） */
    get doc(): MockDocument { return this._activeDoc; }
}
