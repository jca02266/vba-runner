/**
 * 軽量な Excel Worksheet モック
 *
 * 使い方:
 * const mockWs = new MockWorksheet('Sheet1');
 * mockWs.setCellValue('A1', 100);
 * mockWs.setCellValue('B1:B5', [[1], [2], [3], [4], [5]]);
 *
 * 対応アドレス形式:
 * - 'A1'             単一セル
 * - 'A1:B5'          単純範囲
 * - 'A1:C5,E6:F8'    Union（カンマ区切り）— 全サブ範囲を結合
 * - 'A1:C5 B2:D4'    Intersection（スペース区切り）— 重なり部分のみ
 * - '$A$1:$B$5'      絶対参照 — $ を除去して通常参照として扱う
 * - 'A1#'            スピル参照 — # を除去して単一セルとして扱う
 * - 'A1.:.A100'      トリム参照（Excel 365）— .:. / .: / :. を : に正規化して扱う
 * - '@A1:A10'        implicit intersection operator（Excel 365）— @ を除去して通常参照として扱う
 */

import { VbaType, VbaDefaultProperty, vbaEmpty } from '../vba-types';
import type { BuiltinOverload } from '../evaluator';

interface CellRect {
    startCol: number;
    startRow: number;
    endCol: number;
    endRow: number;
}

// ========== フォーマット/スタイルのノーオプスタブ ==========

class MockFont {
    get Bold() { return false; }
    set Bold(_v: any) {}
    get Italic() { return false; }
    set Italic(_v: any) {}
    get Underline() { return 0; }
    set Underline(_v: any) {}
    get Color() { return 0; }
    set Color(_v: any) {}
    get ColorIndex() { return 0; }
    set ColorIndex(_v: any) {}
    get Size() { return 11; }
    set Size(_v: any) {}
    get Name() { return 'Calibri'; }
    set Name(_v: any) {}
}

class MockInterior {
    get Color() { return 0; }
    set Color(_v: any) {}
    get ColorIndex() { return 0; }
    set ColorIndex(_v: any) {}
    get Pattern() { return 0; }
    set Pattern(_v: any) {}
    get PatternColor() { return 0; }
    set PatternColor(_v: any) {}
}

class MockBorder {
    get LineStyle() { return 0; }
    set LineStyle(_v: any) {}
    get Weight() { return 0; }
    set Weight(_v: any) {}
    get Color() { return 0; }
    set Color(_v: any) {}
    get ColorIndex() { return 0; }
    set ColorIndex(_v: any) {}
}

class MockBorders {
    Item() { return new MockBorder(); }
    get LineStyle() { return 0; }
    set LineStyle(_v: any) {}
    get Weight() { return 0; }
    set Weight(_v: any) {}
    get Color() { return 0; }
    set Color(_v: any) {}
}

// ========== 行/列コレクションのスタブ ==========

/** Rows コレクションの最小スタブ。ws.Rows.Count や ws.Rows(n).Hidden などで使う。 */
export class MockRows {
    Count = 1048576;
    get Hidden() { return false; }
    set Hidden(_v: any) {}
    get RowHeight() { return 15; }
    set RowHeight(_v: any) {}
    AutoFit() {}
    Select() {}
    Delete() {}
}

/** Columns コレクションの最小スタブ。ws.Columns.Count や ws.Columns(n).ColumnWidth などで使う。 */
export class MockColumns {
    Count = 16384;
    get Hidden() { return false; }
    set Hidden(_v: any) {}
    get ColumnWidth() { return 8; }
    set ColumnWidth(_v: any) {}
    AutoFit() {}
    Select() {}
    Delete() {}
}

// ========== Range ==========

export class MockRange implements VbaType, VbaDefaultProperty {
    readonly __vbaDefault__ = true as const;
    /** TypeName(r) → "Range"、TypeOf r Is Range → True になるための型メタデータ */
    readonly __vbaTypeName__ = 'Range';

    // フィールド名は `Value` アクセサと大文字小文字で衝突しないよう `_value` にする
    // （case-insensitive なプロパティ解決がアクセサより先にフィールドを拾うのを防ぐ）
    constructor(private _value: any) {}

    get Value(): any {
        return this._value;
    }

    set Value(val: any) {
        this._value = val;
    }

    // JS 数値コンテキスト（+mockRange 等）用。VBA evaluator は Value getter を使うため呼ばれない。
    valueOf(): any {
        return this._value;
    }

    // ==============================
    // 位置情報（Cells で設定される）
    // ==============================

    Row = 1;
    Column = 1;

    get Address(): string {
        return `$A$${this.Row}`;
    }

    // ==============================
    // ナビゲーション
    // ==============================

    /** ws.Cells(lastRow, col).End(xlUp) などで使う。スタブは Row=1 / Column=1 を返す。 */
    End(_direction?: any): MockRange {
        return new MockRange(0);
    }

    get EntireRow(): MockRange {
        return new MockRange(0);
    }

    get EntireColumn(): MockRange {
        return new MockRange(0);
    }

    get Offset(): MockRange {
        return new MockRange(0);
    }

    Resize() { return new MockRange(0); }

    // ==============================
    // スタイル / フォーマット（ノーオプ）
    // ==============================

    get NumberFormat() { return 'General'; }
    set NumberFormat(_v: any) {}

    get NumberFormatLocal() { return 'G/標準'; }
    set NumberFormatLocal(_v: any) {}

    get HorizontalAlignment() { return 0; }
    set HorizontalAlignment(_v: any) {}

    get VerticalAlignment() { return 0; }
    set VerticalAlignment(_v: any) {}

    get WrapText() { return false; }
    set WrapText(_v: any) {}

    get MergeCells() { return false; }
    set MergeCells(_v: any) {}

    get RowHeight() { return 15; }
    set RowHeight(_v: any) {}

    get ColumnWidth() { return 8; }
    set ColumnWidth(_v: any) {}

    get Hidden() { return false; }
    set Hidden(_v: any) {}

    get Locked() { return true; }
    set Locked(_v: any) {}

    get Font(): MockFont { return new MockFont(); }
    get Interior(): MockInterior { return new MockInterior(); }
    get Borders(): MockBorders { return new MockBorders(); }

    // ==============================
    // アクション（ノーオプ）
    // ==============================

    Select() {}
    Activate() {}
    Copy() {}
    Cut() {}
    Paste() {}
    PasteSpecial() {}
    Delete() {}
    Insert() {}
    Clear() {}
    ClearContents() {}
    ClearFormats() {}
    Merge() {}
    UnMerge() {}
    AutoFit() {}
    AutoFill() {}
    Sort() {}
    FillDown() {}
    FillRight() {}
    Find() { return null; }

    // ==============================
    // コレクション
    // ==============================

    Rows() { return new MockRows(); }
    Columns() { return new MockColumns(); }

    get Count() { return 1; }
}

// ========== Worksheet ==========

export class MockWorksheet implements VbaType {
    /** TypeName(ws) → "Worksheet"、TypeOf ws Is Worksheet → True になるための型メタデータ */
    readonly __vbaTypeName__ = 'Worksheet';

    private name: string;
    private cells: Map<string, any> = new Map();

    constructor(name: string) {
        this.name = name;
    }

    /**
     * Range メソッド：
     * - 1引数（Address）: Address 文字列で指定したセル/範囲を返す
     * - 2引数（Cell1, Cell2）: 2つの角セル（Address 文字列または Range オブジェクト）を
     *   結ぶ矩形を返す（VBA の `Range(Cell1, Cell2)` 形式）。VBA 自体には2引数固有の
     *   引数名はないが、ここでは Cell1/Cell2 という名前で `__vbaOverloads__` に登録する。
     */
    Range(cell1: string, cell2?: string | MockRange): MockRange {
        if (cell2 !== undefined) {
            const r1 = this.cellRectFromRangeArg(cell1);
            const r2 = this.cellRectFromRangeArg(cell2);
            if (!r1 || !r2) return new MockRange(0);
            const area: CellRect = {
                startCol: Math.min(r1.startCol, r2.startCol),
                startRow: Math.min(r1.startRow, r2.startRow),
                endCol: Math.max(r1.endCol, r2.endCol),
                endRow: Math.max(r1.endRow, r2.endRow),
            };
            return this.rangeFromArea(area);
        }

        const areas = this.resolveAddress(cell1);

        if (areas.length === 0) {
            return new MockRange(0);
        }

        if (areas.length === 1) {
            return this.rangeFromArea(areas[0]);
        }

        // 複数エリア（Union）: 全エリアの行を縦に結合して返す
        const self = this;
        const combined: any[][] = [];
        for (const area of areas) {
            combined.push(...this.getRectValues(area));
        }
        const unionRange = new MockRange(combined);
        Object.defineProperty(unionRange, 'Value', {
            get: () => {
                const result: any[][] = [];
                for (const area of areas) {
                    result.push(...self.getRectValues(area));
                }
                return result;
            },
            set: (val: any) => {
                for (const area of areas) {
                    self.setRectValue(area, val);
                }
            },
        });
        return unionRange;
    }

    /** `CellRect` から `Value` の get/set を `cells` と同期させた `MockRange` を作る */
    private rangeFromArea(area: CellRect): MockRange {
        const self = this;
        if (area.startCol === area.endCol && area.startRow === area.endRow) {
            // 単一セル — setter で cells を同期
            const cellAddr = this.cellAddress(area.startCol, area.startRow);
            const range = new MockRange(this.cells.get(cellAddr) ?? vbaEmpty);
            range.Row = area.startRow;
            range.Column = area.startCol;
            Object.defineProperty(range, 'Value', {
                get: () => self.cells.get(cellAddr) ?? vbaEmpty,
                set: (val: any) => { self.cells.set(cellAddr, val); },
            });
            return range;
        }
        // 複数セル範囲 — setter で cells に書き戻す
        const range = new MockRange(this.getRectValues(area));
        range.Row = area.startRow;
        range.Column = area.startCol;
        Object.defineProperty(range, 'Value', {
            get: () => self.getRectValues(area),
            set: (val: any) => { self.setRectValue(area, val); },
        });
        return range;
    }

    /** `Range(Cell1, Cell2)` の角セル引数（Address 文字列 or Range オブジェクト）を `CellRect` に解決する */
    private cellRectFromRangeArg(arg: string | MockRange): CellRect | null {
        if (arg instanceof MockRange) {
            return { startCol: arg.Column, startRow: arg.Row, endCol: arg.Column, endRow: arg.Row };
        }
        const normalized = String(arg).toUpperCase().replace(/\$/g, '').trim();
        return this.parseSingleRange(normalized);
    }

    /**
     * Cells メソッド：行・列インデックス（1始まり）で単一セルを返す
     * 引数なし（ws.Cells）で呼ばれた場合はシート全体を表すスタブを返す。
     * 注: length=0 にするため引数を宣言しない。VBA の ws.Cells(r, c) 呼び出しは
     * evaluator が引数を渡して呼ぶため動作に問題はない。
     */
    Cells(...args: any[]): MockRange {
        const [row, col] = args;
        if (row === undefined || col === undefined) {
            // ws.Cells（引数なし）: シート全体を表すスタブ
            return new MockRange(0);
        }
        const r = Number(row);
        const c = Number(col);
        let colStr = '';
        let cv = c;
        while (cv > 0) {
            colStr = String.fromCharCode(((cv - 1) % 26) + 65) + colStr;
            cv = Math.floor((cv - 1) / 26);
        }
        const range = this.Range(`${colStr}${r}`);
        range.Row = r;
        range.Column = c;
        return range;
    }

    /**
     * セルに値を設定
     * @param address セルアドレス ('A1') または範囲 ('A1:A5', 'A1:C5,E6:F8' など)
     * @param value スカラー値または配列
     */
    setCellValue(address: string, value: any): void {
        for (const area of this.resolveAddress(address)) {
            this.setRectValue(area, value);
        }
    }

    /**
     * セルから値を取得
     */
    getCellValue(address: string): any {
        return this.cells.get(address.toUpperCase().replace(/\$/g, '')) ?? vbaEmpty;
    }

    /**
     * ワークシート内の全セルをダンプ（デバッグ用）
     */
    dump(): Record<string, any> {
        const result: Record<string, any> = {};
        this.cells.forEach((value, key) => { result[key] = value; });
        return result;
    }

    // ==============================
    // 行/列コレクション
    // ==============================

    /**
     * ws.Rows.Count / ws.Rows(n).Hidden などで使う。
     * length=0 のメソッドにすることで ws.Rows（引数なし）でも auto-call される。
     */
    Rows(..._args: any[]): MockRows {
        return new MockRows();
    }

    /**
     * ws.Columns.Count / ws.Columns(n).ColumnWidth などで使う。
     */
    Columns(..._args: any[]): MockColumns {
        return new MockColumns();
    }

    /** ws.UsedRange: 使用済みセル範囲のスタブ（A1 を返す） */
    get UsedRange(): MockRange {
        return this.Range('A1');
    }

    // ==============================
    // プロパティ
    // ==============================

    get Name(): string {
        return this.name;
    }

    get Index(): number { return 1; }

    get Visible(): boolean { return true; }
    set Visible(_v: any) {}

    get EnableSelection(): number { return 0; }
    set EnableSelection(_v: any) {}

    // ==============================
    // アクション（ノーオプ）
    // ==============================

    Activate() {}
    Select() {}
    Copy() {}
    Move() {}
    Delete() {}
    Protect() {}
    Unprotect() {}

    // ==============================
    // アドレス解析
    // ==============================

    /**
     * アドレス文字列を CellRect の配列に解決する。
     * Union（カンマ）→ 各サブ範囲を個別に返す。
     * Intersection（スペース）→ 全サブ範囲の重なり矩形を返す（重なりなしは空配列）。
     */
    private resolveAddress(address: string): CellRect[] {
        // 正規化: 大文字化 / $ 除去 / # 除去（スピル）/ @ 除去（implicit intersection）
        // / トリム参照 A1.:.A100 → A1:A100（.? : .? のドットを除去）/ 前後空白除去
        const normalized = address
            .toUpperCase()
            .replace(/\$/g, '')
            .replace(/#/g, '')
            .replace(/^@/, '')
            .replace(/\.?:\.?/g, ':')
            .trim();

        if (normalized.length === 0) return [];

        // Union: カンマで分割
        if (normalized.includes(',')) {
            return normalized
                .split(',')
                .map(p => this.parseSingleRange(p.trim()))
                .filter((r): r is CellRect => r !== null);
        }

        // Intersection or single: スペースで分割
        const parts = normalized.split(/\s+/).filter(p => p.length > 0);
        if (parts.length === 1) {
            const r = this.parseSingleRange(parts[0]);
            return r ? [r] : [];
        }

        // Intersection: 全パーツの重なり矩形を計算
        const rects = parts.map(p => this.parseSingleRange(p)).filter((r): r is CellRect => r !== null);
        const intersection = this.computeIntersection(rects);
        return intersection ? [intersection] : [];
    }

    private parseSingleRange(part: string): CellRect | null {
        if (!part || part.length === 0) return null;
        if (part.includes(':')) {
            const [start, end] = part.split(':');
            const sc = this.parseAddress(start);
            const ec = this.parseAddress(end);
            if (!sc || !ec) return null;
            return { startCol: sc[0], startRow: sc[1], endCol: ec[0], endRow: ec[1] };
        }
        const c = this.parseAddress(part);
        if (!c) return null;
        return { startCol: c[0], startRow: c[1], endCol: c[0], endRow: c[1] };
    }

    private computeIntersection(rects: CellRect[]): CellRect | null {
        if (rects.length === 0) return null;
        let result = rects[0];
        for (let i = 1; i < rects.length; i++) {
            const r = rects[i];
            const startCol = Math.max(result.startCol, r.startCol);
            const startRow = Math.max(result.startRow, r.startRow);
            const endCol   = Math.min(result.endCol,   r.endCol);
            const endRow   = Math.min(result.endRow,   r.endRow);
            if (startCol > endCol || startRow > endRow) return null;
            result = { startCol, startRow, endCol, endRow };
        }
        return result;
    }

    // ==============================
    // セル読み書きヘルパー
    // ==============================

    private getRectValues(area: CellRect): any[][] {
        const result: any[][] = [];
        for (let r = area.startRow; r <= area.endRow; r++) {
            const row: any[] = [];
            for (let c = area.startCol; c <= area.endCol; c++) {
                row.push(this.cells.get(this.cellAddress(c, r)) ?? vbaEmpty);
            }
            result.push(row);
        }
        return result;
    }

    private setRectValue(area: CellRect, value: any): void {
        if (Array.isArray(value)) {
            const is1D = value.length === 0 || !Array.isArray(value[0]);
            if (is1D) {
                // 1D 配列: 1行分の列値として全行に繰り返し適用（VBA の Array() 代入と同じ）
                for (let r = area.startRow; r <= area.endRow; r++) {
                    let ci = 0;
                    for (let c = area.startCol; c <= area.endCol; c++) {
                        this.cells.set(this.cellAddress(c, r), value[ci] ?? '');
                        ci++;
                    }
                }
            } else {
                let ri = 0;
                for (let r = area.startRow; r <= area.endRow; r++) {
                    const rowData = value[ri];
                    if (Array.isArray(rowData)) {
                        let ci = 0;
                        for (let c = area.startCol; c <= area.endCol; c++) {
                            this.cells.set(this.cellAddress(c, r), rowData[ci] ?? '');
                            ci++;
                        }
                    }
                    ri++;
                }
            }
        } else {
            for (let r = area.startRow; r <= area.endRow; r++) {
                for (let c = area.startCol; c <= area.endCol; c++) {
                    this.cells.set(this.cellAddress(c, r), value);
                }
            }
        }
    }

    // ==============================
    // 基本ヘルパー
    // ==============================

    private parseAddress(address: string): [number, number] | null {
        const match = address.match(/^([A-Z]+)(\d+)$/);
        if (!match) return null;
        return [this.columnToNumber(match[1]), parseInt(match[2], 10)];
    }

    private columnToNumber(col: string): number {
        let result = 0;
        for (let i = 0; i < col.length; i++) {
            result = result * 26 + (col.charCodeAt(i) - 64);
        }
        return result;
    }

    private cellAddress(col: number, row: number): string {
        let colStr = '';
        let c = col;
        while (c > 0) {
            colStr = String.fromCharCode(((c - 1) % 26) + 65) + colStr;
            c = Math.floor((c - 1) / 26);
        }
        return `${colStr}${row}`;
    }
}

// VBA 自体にはない「引数の個数で意味が変わる」組み込み関数専用のオーバーロード機構
// （`Evaluator.registerOverloadedBuiltin` 参照）を、モックオブジェクトのメソッドにも適用する。
// `obj.Range(...)` 呼び出し（Tier 6 の defaultBindingObject 経由も含む）は `resolveCallArgs`
// がこのプロパティを見て引数数検証・名前付き引数解決を行う。本体（Range 自身）は変更しない。
(MockWorksheet.prototype.Range as any).__vbaOverloads__ = [
    { params: [{ name: 'Cell1' }] },
    { params: [{ name: 'Cell1' }, { name: 'Cell2' }] },
] satisfies BuiltinOverload[];

