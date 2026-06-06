/**
 * Access Application モック
 *
 * `injectAccessStub(evaluator)` と組み合わせて使うと、Access 依存コードを
 * エンジン上で実行できる。
 *
 * 対応する VBA パターン（代表例）:
 *   Set db = CurrentDb()
 *   Set rs = db.OpenRecordset("TableName")
 *   Do While Not rs.EOF
 *       Debug.Print rs.Fields("Name").Value
 *       rs.MoveNext
 *   Loop
 *   DoCmd.SetWarnings False
 *   DoCmd.RunSQL "INSERT INTO ..."
 *   Forms("frmCustomer").txtName.Value
 *
 * フォームコントロール（Me.txtName 等）は setControl() で事前登録した分のみ読み書き可能。
 * 登録されていないコントロール名へのアクセスはエラーになる（VBA Runner の制約）。
 */

import { VbaType } from '../vba-types';

// ========== DAO: Field ==========

export class MockField implements VbaType {
    readonly __vbaTypeName__ = 'Field';
    Name: string;
    Value: any;
    Type: number = 10;   // dbText
    Size: number = 255;

    constructor(name: string, value: any = '') {
        this.Name = name;
        this.Value = value;
    }
}

// ========== DAO: Fields コレクション ==========

/**
 * Recordset.Fields コレクション。
 * rs.Fields("FieldName").Value / rs.Fields(0).Value の両形式に対応。
 *
 * VBA の `rs.Fields("name")` は CallExpression として評価されるため、
 * Fields は関数として返す必要がある。Count 等のプロパティも関数に付与する。
 */
class MockFieldsCollection {
    private _store: Map<string, MockField> = new Map();

    /** フィールドを追加（Recordset 内部から呼ぶ） */
    _add(name: string, value: any = ''): void {
        this._store.set(name.toLowerCase(), new MockField(name, value));
    }

    /** フィールドの値をまとめて同期する（行移動時） */
    _syncRow(row: Record<string, any>): void {
        for (const [k, v] of Object.entries(row)) {
            const key = k.toLowerCase();
            if (this._store.has(key)) {
                this._store.get(key)!.Value = v;
            } else {
                this._store.set(key, new MockField(k, v));
            }
        }
    }

    get Count(): number { return this._store.size; }

    Item(nameOrIndex: string | number): MockField {
        if (typeof nameOrIndex === 'number') {
            const values = [...this._store.values()];
            return values[nameOrIndex] ?? new MockField(String(nameOrIndex));
        }
        const key = nameOrIndex.toLowerCase();
        if (!this._store.has(key)) {
            this._store.set(key, new MockField(nameOrIndex));
        }
        return this._store.get(key)!;
    }
}

/**
 * evaluator から `rs.Fields("name")` と呼ばれるために関数として返す。
 * `rs.Fields.Count` のプロパティアクセスも機能するよう Count / Item を付与する。
 */
function makeFieldsAccessor(col: MockFieldsCollection): ((nameOrIndex: string | number) => MockField) & { Count: number; Item: (n: string | number) => MockField } {
    const fn = (nameOrIndex: string | number) => col.Item(nameOrIndex);
    Object.defineProperty(fn, 'Count', { get: () => col.Count, enumerable: true });
    fn.Item = (n: string | number) => col.Item(n);
    return fn as any;
}

// ========== DAO: Recordset ==========

export class MockRecordset implements VbaType {
    readonly __vbaTypeName__ = 'Recordset';

    private _rows: Record<string, any>[];
    private _index: number = 0;
    private _col: MockFieldsCollection;

    /** rs.Fields("name") / rs.Fields(0) / rs.Fields.Count に対応 */
    readonly Fields: ReturnType<typeof makeFieldsAccessor>;

    NoMatch: boolean = true;

    constructor(rows: Record<string, any>[] = []) {
        this._rows = rows;
        this._col = new MockFieldsCollection();
        this.Fields = makeFieldsAccessor(this._col);
        this._syncFields();
    }

    private _syncFields(): void {
        if (this._rows.length > 0 && this._index < this._rows.length) {
            this._col._syncRow(this._rows[this._index]);
        }
    }

    get EOF(): boolean { return this._index >= this._rows.length; }
    get BOF(): boolean { return this._rows.length === 0 || this._index < 0; }
    get RecordCount(): number { return this._rows.length; }
    get AbsolutePosition(): number { return this._index; }
    set AbsolutePosition(v: number) { this._index = v; this._syncFields(); }

    MoveNext() { this._index++; this._syncFields(); }
    MovePrevious() { this._index = Math.max(-1, this._index - 1); this._syncFields(); }
    MoveFirst() { this._index = 0; this._syncFields(); }
    MoveLast() { this._index = Math.max(0, this._rows.length - 1); this._syncFields(); }

    AddNew() {}
    Edit() {}
    Update() {}
    Delete() {}
    Close() {}

    FindFirst(_criteria: string) { this.NoMatch = true; }
    FindNext(_criteria: string) { this.NoMatch = true; }
    FindPrevious(_criteria: string) { this.NoMatch = true; }
    FindLast(_criteria: string) { this.NoMatch = true; }
}

// ========== DAO: Database ==========

export class MockDatabase implements VbaType {
    readonly __vbaTypeName__ = 'Database';

    private _tables: Map<string, Record<string, any>[]> = new Map();

    OpenRecordset(_nameOrSql: string): MockRecordset {
        const key = _nameOrSql.toLowerCase();
        return new MockRecordset(this._tables.get(key) ?? []);
    }

    Execute(_sql: string, _options?: number) {}
    Close() {}

    /** テスト用: テーブルまたはクエリのデータを事前設定する */
    setTableData(name: string, rows: Record<string, any>[]): void {
        this._tables.set(name.toLowerCase(), rows);
    }
}

// ========== DoCmd ==========

/** DoCmd.OpenForm / DoCmd.RunSQL 等を全てノーオプにするスタブ */
export class MockDoCmd {
    OpenForm() {}
    OpenReport() {}
    OpenQuery() {}
    OpenTable() {}
    Close() {}
    RunSQL() {}
    SetWarnings() {}
    Hourglass() {}
    RunMacro() {}
    RunCommand() {}
    GoToRecord() {}
    FindRecord() {}
    ApplyFilter() {}
    RemoveFilterSort() {}
    ShowAllRecords() {}
    PrintOut() {}
    OutputTo() {}
    TransferDatabase() {}
    TransferSpreadsheet() {}
    TransferText() {}
    SendObject() {}
    Save() {}
    Quit() {}
    CancelEvent() {}
    Echo() {}
    Maximize() {}
    Minimize() {}
    Restore() {}
    MoveSize() {}
    SelectObject() {}
    SendKeys() {}
    AddMenu() {}
    Beep() {}
    DeleteObject() {}
    CopyObject() {}
    RenameObject() {}
    RepaintObject() {}
    Requery() {}
}

// ========== フォームコントロール ==========

export class MockControl implements VbaType {
    readonly __vbaTypeName__ = 'Control';
    Name: string;
    Value: any = '';
    Text: string = '';
    Visible: boolean = true;
    Enabled: boolean = true;
    Locked: boolean = false;
    BackColor: number = 16777215;
    ForeColor: number = 0;
    FontBold: boolean = false;
    FontItalic: boolean = false;
    Caption: string = '';
    Tag: string = '';
    TabIndex: number = 0;

    constructor(name: string, value: any = '') {
        this.Name = name;
        this.Caption = name;
        this.Value = value;
    }

    SetFocus() {}
    Requery() {}
}

// ========== フォーム ==========

export class MockForm implements VbaType {
    readonly __vbaTypeName__ = 'Form';
    Name: string;
    Caption: string;
    Visible: boolean = true;
    Modal: boolean = false;
    RecordSource: string = '';
    Filter: string = '';
    FilterOn: boolean = false;
    AllowEdits: boolean = true;
    AllowAdditions: boolean = true;
    AllowDeletions: boolean = true;
    Dirty: boolean = false;
    Tag: string = '';

    constructor(name: string = 'MockForm') {
        this.Name = name;
        this.Caption = name;
    }

    /**
     * コントロールを事前登録する。
     * 登録した名前は evaluator の resolveObjectMemberKey で発見可能になるため、
     * VBA コードから `frm.txtName.Value` のようにアクセスできる。
     *
     * @example
     *   const form = app.Forms('frmCustomer');
     *   form.setControl('txtName', 'Alice');
     *   form.setControl('txtAge', 30);
     *   runner.set('Me', form);
     */
    setControl(name: string, value: any = ''): MockControl {
        const ctrl = new MockControl(name, value);
        // 大文字小文字両方で自インスタンスに直接プロパティとして登録
        Object.defineProperty(this, name, {
            get: () => ctrl,
            set: (v: any) => { ctrl.Value = v; },
            enumerable: true,
            configurable: true,
        });
        if (name !== name.toLowerCase()) {
            Object.defineProperty(this, name.toLowerCase(), {
                get: () => ctrl,
                set: (v: any) => { ctrl.Value = v; },
                enumerable: true,
                configurable: true,
            });
        }
        return ctrl;
    }

    Refresh() {}
    Requery() {}
    Repaint() {}
    SetFocus() {}
    Close() {}
    Undo() {}
}

// ========== Access Application ==========

export class MockAccessApplication {
    private _db: MockDatabase = new MockDatabase();
    private _doCmd: MockDoCmd = new MockDoCmd();
    private _forms: Map<string, MockForm> = new Map();

    // ==============================
    // Tier 6 グローバルメンバー
    // ==============================

    /** VBA: Set db = CurrentDb() */
    CurrentDb(): MockDatabase { return this._db; }

    /** VBA: DoCmd.OpenForm "frmName" */
    get DoCmd(): MockDoCmd { return this._doCmd; }

    /**
     * VBA: Forms("frmName").txtField.Value
     * 存在しないフォーム名は自動作成する。
     */
    Forms(nameOrIndex: string | number = ''): MockForm {
        const name = String(nameOrIndex);
        const key = name.toLowerCase();
        if (!this._forms.has(key)) {
            this._forms.set(key, new MockForm(name));
        }
        return this._forms.get(key)!;
    }

    // ==============================
    // Application レベルの設定（ノーオプ）
    // ==============================

    get Name(): string { return 'Microsoft Access'; }
    get Version(): string { return '16.0'; }
    get Application(): this { return this; }

    get ScreenUpdating() { return true; }
    set ScreenUpdating(_v: any) {}

    get DisplayAlerts() { return true; }
    set DisplayAlerts(_v: any) {}

    /** VBA: Application.Echo False / Application.Echo True */
    Echo() {}

    Quit() {}

    // ==============================
    // テスト用ユーティリティ
    // ==============================

    /** CurrentDb() が返す MockDatabase への参照（テーブルデータを仕込む用） */
    get db(): MockDatabase { return this._db; }
}
