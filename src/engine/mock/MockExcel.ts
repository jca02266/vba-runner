/**
 * Excel Application モック
 *
 * `setDefaultBindingObject(new MockApplication())` と組み合わせて使うと、
 * VBA コード内の修飾なし Excel グローバル呼び出しをエンジン上で動作させられる。
 *
 * ```typescript
 * import { MockApplication } from '../src/engine/mock/MockExcel';
 *
 * const app = new MockApplication();
 * app.ActiveSheet.setCellValue('A1', 42);
 * ev.setDefaultBindingObject(app);
 *
 * // VBA: Range("A1").Value  → 42
 * // VBA: Cells(1, 1).Value → 42
 * // VBA: ActiveSheet.Name  → "Sheet1"
 * ```
 *
 * MockApplication は opt-in。`setDefaultBindingObject` を呼ばなければ既存の動作と変わらない。
 */

import { MockWorksheet, MockRange } from './MockWorksheet';

export { MockWorksheet, MockRange };

/** ActiveWorkbook / ThisWorkbook の最小スタブ */
export class MockWorkbook {
    constructor(
        public readonly Name: string = 'MockWorkbook.xlsx',
        public readonly Path: string = '',
    ) {}

    get FullName(): string {
        return this.Path ? `${this.Path}\\${this.Name}` : this.Name;
    }
}

/**
 * Excel Application モック
 *
 * Tier 6 で解決される Excel グローバル関数・プロパティを提供する。
 * VBA コードが `Range("A1")` や `ActiveSheet.Name` のように修飾なしで呼ぶ場合、
 * `evaluator.setDefaultBindingObject(app)` によりこのオブジェクトが参照される。
 */
export class MockApplication {
    private _sheets: Map<string, MockWorksheet> = new Map();
    private _activeSheetName: string = 'Sheet1';
    private _workbook: MockWorkbook;

    constructor(workbookName: string = 'MockWorkbook.xlsx') {
        this._workbook = new MockWorkbook(workbookName);
        this._sheets.set('Sheet1', new MockWorksheet('Sheet1'));
    }

    // ==============================
    // Tier 6 グローバルメンバー
    // ==============================

    /**
     * 現在のアクティブシートを返す。
     * VBA: `ActiveSheet.Range("A1")` / `Set ws = ActiveSheet`
     */
    get ActiveSheet(): MockWorksheet {
        return this.Sheets(this._activeSheetName);
    }

    /**
     * シートを名前またはインデックス（1始まり）で取得する。存在しない場合は自動作成。
     * VBA: `Sheets("Sheet1").Range("A1")` / `Sheets(1).Name`
     */
    Sheets(nameOrIndex: string | number): MockWorksheet {
        const name =
            typeof nameOrIndex === 'string' ? nameOrIndex : `Sheet${nameOrIndex}`;
        if (!this._sheets.has(name)) {
            this._sheets.set(name, new MockWorksheet(name));
        }
        return this._sheets.get(name)!;
    }

    /**
     * アクティブシートのセル範囲を返す。
     * VBA: `Range("A1:B5").Value` / `Set r = Range("A1")`
     */
    Range(address: string): MockRange {
        return this.ActiveSheet.Range(address);
    }

    /**
     * アクティブシートの単一セルを行・列インデックスで返す（1始まり）。
     * VBA: `Cells(1, 1).Value = 100`
     */
    Cells(row: number, col: number): MockRange {
        return this.ActiveSheet.Cells(row, col);
    }

    /** VBA: `ActiveWorkbook.Name` */
    get ActiveWorkbook(): MockWorkbook {
        return this._workbook;
    }

    /** VBA: `ThisWorkbook.Name` */
    get ThisWorkbook(): MockWorkbook {
        return this._workbook;
    }

    /** VBA: `Application.Name` */
    get Name(): string {
        return 'Microsoft Excel';
    }

    // ==============================
    // テスト用ユーティリティ
    // ==============================

    /**
     * アクティブシートを切り替える。
     * 存在しないシート名を指定すると自動作成する。
     */
    setActiveSheet(nameOrIndex: string | number): void {
        const name =
            typeof nameOrIndex === 'string' ? nameOrIndex : `Sheet${nameOrIndex}`;
        this._activeSheetName = name;
        if (!this._sheets.has(name)) {
            this._sheets.set(name, new MockWorksheet(name));
        }
    }

    /** 登録されているシート名の一覧を返す */
    listSheets(): string[] {
        return Array.from(this._sheets.keys());
    }

    /** 全シートのセルデータをクリアし、Sheet1 だけの初期状態に戻す */
    clear(): void {
        this._sheets.clear();
        this._sheets.set('Sheet1', new MockWorksheet('Sheet1'));
        this._activeSheetName = 'Sheet1';
    }
}
