/**
 * exceljs バックエンドを使った Excel アダプター
 *
 * MockApplication と同じインターフェースを実装し、VBARunner の excelStub として渡すことで
 * VBA の実行結果を実際の .xlsx ファイルに反映させる。
 *
 * Usage:
 *   const wb = new ExcelJS.Workbook();
 *   await wb.xlsx.readFile('input.xlsx');
 *   const app = new ExcelJsApplication(wb);
 *   const runner = new VBARunner(['macro.bas'], { excelStub: app as any });
 *   runner.run('Main');
 *   await wb.xlsx.writeFile('output.xlsx');
 */

import ExcelJS from 'exceljs';
import { VbaType, VbaDefaultProperty } from '../../src/engine/vba-types';
import type { BuiltinOverload } from '../../src/engine/evaluator';

// ---------- セルアドレス変換ユーティリティ ----------

function colLetterToNum(letters: string): number {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
}

function numToColLetter(n: number): string {
    let s = '';
    while (n > 0) {
        s = String.fromCharCode(65 + (n - 1) % 26) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function parseCell(addr: string): { row: number; col: number } | null {
    const m = addr.replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return null;
    return { col: colLetterToNum(m[1]), row: parseInt(m[2]) };
}

function parseRange(addr: string): { sr: number; sc: number; er: number; ec: number } | null {
    const plain = addr.includes('!') ? addr.slice(addr.indexOf('!') + 1) : addr;
    const colon = plain.indexOf(':');
    if (colon < 0) {
        const c = parseCell(plain);
        if (!c) return null;
        return { sr: c.row, sc: c.col, er: c.row, ec: c.col };
    }
    const c1 = parseCell(plain.slice(0, colon));
    const c2 = parseCell(plain.slice(colon + 1));
    if (!c1 || !c2) return null;
    return {
        sr: Math.min(c1.row, c2.row), sc: Math.min(c1.col, c2.col),
        er: Math.max(c1.row, c2.row), ec: Math.max(c1.col, c2.col),
    };
}

/** exceljs セルの値を VBA 互換の形式で返す */
function readCell(cell: ExcelJS.Cell): any {
    const v = cell.value;
    if (v === null || v === undefined) return '';
    // 数式セル — キャッシュされた計算結果を返す
    if (typeof v === 'object' && 'result' in v) return (v as ExcelJS.CellFormulaValue).result ?? 0;
    // リッチテキスト — プレーンテキストに変換
    if (typeof v === 'object' && 'richText' in v)
        return (v as ExcelJS.CellRichTextValue).richText.map(r => r.text).join('');
    return v;
}

// ---------- ExcelJsRange ----------

export class ExcelJsRange implements VbaType, VbaDefaultProperty {
    readonly __vbaDefault__ = true as const;
    readonly __vbaTypeName__ = 'Range';

    constructor(
        private _ws: ExcelJS.Worksheet,
        public readonly startRow: number,
        public readonly startCol: number,
        public readonly endRow: number,
        public readonly endCol: number,
    ) {}

    get Row() { return this.startRow; }
    get Column() { return this.startCol; }

    get Address(): string {
        const a = `$${numToColLetter(this.startCol)}$${this.startRow}`;
        if (this.startRow === this.endRow && this.startCol === this.endCol) return a;
        return `${a}:$${numToColLetter(this.endCol)}$${this.endRow}`;
    }

    // --------- Value 読み書き（コアメソッド） ---------

    get Value(): any {
        if (this.startRow === this.endRow && this.startCol === this.endCol) {
            return readCell(this._ws.getCell(this.startRow, this.startCol));
        }
        const rows: any[][] = [];
        for (let r = this.startRow; r <= this.endRow; r++) {
            const row: any[] = [];
            for (let c = this.startCol; c <= this.endCol; c++) {
                row.push(readCell(this._ws.getCell(r, c)));
            }
            rows.push(row);
        }
        return rows;
    }

    set Value(val: any) {
        if (Array.isArray(val) && Array.isArray(val[0])) {
            // 2D 配列の一括代入
            for (let ri = 0; ri < (val as any[][]).length; ri++) {
                for (let ci = 0; ci < (val[ri] as any[]).length; ci++) {
                    this._ws.getCell(this.startRow + ri, this.startCol + ci).value = val[ri][ci];
                }
            }
        } else {
            // スカラー値を範囲内全セルに代入
            for (let r = this.startRow; r <= this.endRow; r++) {
                for (let c = this.startCol; c <= this.endCol; c++) {
                    this._ws.getCell(r, c).value = val === '' ? null : val;
                }
            }
        }
    }

    // --------- 数式 ---------

    get Formula(): string {
        const cell = this._ws.getCell(this.startRow, this.startCol);
        const v = cell.value;
        if (typeof v === 'object' && v !== null && 'formula' in v) return '=' + (v as ExcelJS.CellFormulaValue).formula;
        return '';
    }

    set Formula(f: string) {
        const formula = f.startsWith('=') ? f.slice(1) : f;
        this._ws.getCell(this.startRow, this.startCol).value = { formula } as ExcelJS.CellFormulaValue;
    }

    // --------- 書式 ---------

    get NumberFormat(): string {
        return (this._ws.getCell(this.startRow, this.startCol).numFmt as string) ?? 'General';
    }
    set NumberFormat(fmt: string) {
        for (let r = this.startRow; r <= this.endRow; r++)
            for (let c = this.startCol; c <= this.endCol; c++)
                this._ws.getCell(r, c).numFmt = fmt;
    }
    get NumberFormatLocal() { return this.NumberFormat; }
    set NumberFormatLocal(v: string) { this.NumberFormat = v; }

    // Font は先頭セルへの実際の書き込みのみ対応（範囲一括は将来の課題）
    get Font() {
        const cell = this._ws.getCell(this.startRow, this.startCol);
        const ws = this._ws;
        const sr = this.startRow, sc = this.startCol, er = this.endRow, ec = this.endCol;
        return {
            get Bold() { return cell.font?.bold ?? false; },
            set Bold(v: boolean) {
                for (let r = sr; r <= er; r++)
                    for (let c = sc; c <= ec; c++) {
                        const tgt = ws.getCell(r, c);
                        tgt.font = { ...tgt.font, bold: v };
                    }
            },
            get Italic() { return cell.font?.italic ?? false; },
            set Italic(v: boolean) {
                for (let r = sr; r <= er; r++)
                    for (let c = sc; c <= ec; c++) {
                        const tgt = ws.getCell(r, c);
                        tgt.font = { ...tgt.font, italic: v };
                    }
            },
            get Size() { return cell.font?.size ?? 11; },
            set Size(v: number) {
                for (let r = sr; r <= er; r++)
                    for (let c = sc; c <= ec; c++) {
                        const tgt = ws.getCell(r, c);
                        tgt.font = { ...tgt.font, size: v };
                    }
            },
            get Name() { return cell.font?.name ?? 'Calibri'; },
            set Name(v: string) {
                for (let r = sr; r <= er; r++)
                    for (let c = sc; c <= ec; c++) {
                        const tgt = ws.getCell(r, c);
                        tgt.font = { ...tgt.font, name: v };
                    }
            },
            get Color() { return 0; },
            set Color(_v: any) {},
            get ColorIndex() { return 0; },
            set ColorIndex(_v: any) {},
            get Underline() { return 0; },
            set Underline(_v: any) {},
        };
    }

    // Interior（背景色）— stub
    get Interior() {
        return { Color: 0, ColorIndex: 0, Pattern: 0, PatternColor: 0 };
    }

    // Borders — stub
    get Borders() {
        const border = { LineStyle: 0, Weight: 0, Color: 0, ColorIndex: 0 };
        return { Item: () => border, LineStyle: 0, Weight: 0, Color: 0 };
    }

    // --------- ナビゲーション ---------

    Offset(rowOff = 0, colOff = 0): ExcelJsRange {
        return new ExcelJsRange(
            this._ws,
            this.startRow + rowOff, this.startCol + colOff,
            this.endRow + rowOff, this.endCol + colOff,
        );
    }

    Resize(rows?: number, cols?: number): ExcelJsRange {
        return new ExcelJsRange(
            this._ws,
            this.startRow, this.startCol,
            this.startRow + (rows ?? 1) - 1,
            this.startCol + (cols ?? 1) - 1,
        );
    }

    /** ws.Cells(lastRow, col).End(xlUp) パターンをサポート */
    End(direction?: number): ExcelJsRange {
        const xlUp = -4162, xlDown = -4121;
        if (direction === xlUp) {
            for (let r = this.startRow; r >= 1; r--) {
                if (readCell(this._ws.getCell(r, this.startCol)) !== '') {
                    return new ExcelJsRange(this._ws, r, this.startCol, r, this.startCol);
                }
            }
            return new ExcelJsRange(this._ws, 1, this.startCol, 1, this.startCol);
        }
        if (direction === xlDown) {
            const maxRow = this._ws.actualRowCount || 65536;
            for (let r = this.startRow; r <= maxRow; r++) {
                if (readCell(this._ws.getCell(r, this.startCol)) === '') {
                    return new ExcelJsRange(this._ws, r - 1, this.startCol, r - 1, this.startCol);
                }
            }
            return new ExcelJsRange(this._ws, maxRow, this.startCol, maxRow, this.startCol);
        }
        return this;
    }

    get EntireRow(): ExcelJsRange {
        return new ExcelJsRange(this._ws, this.startRow, 1, this.endRow, 16384);
    }
    get EntireColumn(): ExcelJsRange {
        return new ExcelJsRange(this._ws, 1, this.startCol, 1048576, this.endCol);
    }

    // --------- コレクション/カウント ---------

    get Count(): number {
        return (this.endRow - this.startRow + 1) * (this.endCol - this.startCol + 1);
    }

    // n に = undefined を使うことで Function.length = 0 になり、
    // エバリュエーターが ws.Rows / ws.Columns（引数なし）を自動呼び出しする。
    Rows(n: number | undefined = undefined): ExcelJsRange | { Count: number } {
        if (n !== undefined)
            return new ExcelJsRange(this._ws, this.startRow + n - 1, this.startCol, this.startRow + n - 1, this.endCol);
        return { Count: this.endRow - this.startRow + 1 };
    }
    Columns(n: number | undefined = undefined): ExcelJsRange | { Count: number } {
        if (n !== undefined)
            return new ExcelJsRange(this._ws, this.startRow, this.startCol + n - 1, this.endRow, this.startCol + n - 1);
        return { Count: this.endCol - this.startCol + 1 };
    }

    // --------- スタイルスタブ ---------

    get HorizontalAlignment() { return 0; }
    set HorizontalAlignment(_v: any) {}
    get VerticalAlignment() { return 0; }
    set VerticalAlignment(_v: any) {}
    get WrapText() { return false; }
    set WrapText(_v: any) {}
    get MergeCells() { return false; }
    set MergeCells(_v: any) {}
    get Hidden() { return false; }
    set Hidden(_v: any) {}
    get Locked() { return true; }
    set Locked(_v: any) {}
    get RowHeight() { return 15; }
    set RowHeight(_v: any) {}
    get ColumnWidth() { return 8; }
    set ColumnWidth(_v: any) {}

    // --------- アクションスタブ ---------

    Select() {}
    Activate() {}
    Copy() {}
    Cut() {}
    Paste() {}
    PasteSpecial() {}
    Delete() {}
    Insert() {}
    Clear() { this.Value = ''; }
    ClearContents() { this.Value = ''; }
    ClearFormats() {}
    Merge() {}
    UnMerge() {}
    AutoFit() {}
    AutoFill() {}
    Sort() {}
    FillDown() {}
    FillRight() {}
    Find() { return null; }
}

// ---------- ExcelJsWorksheet ----------

export class ExcelJsWorksheet implements VbaType {
    readonly __vbaTypeName__ = 'Worksheet';

    constructor(private _ws: ExcelJS.Worksheet) {}

    get Name(): string { return this._ws.name; }

    Cells(row: number, col: number): ExcelJsRange {
        return new ExcelJsRange(this._ws, row, col, row, col);
    }

    Range(cell1: string, cell2?: string | ExcelJsRange): ExcelJsRange {
        if (cell2 !== undefined) {
            const addr2 = typeof cell2 === 'string' ? cell2 : cell2.Address;
            const r1 = parseRange(cell1);
            const r2 = parseRange(addr2);
            if (!r1 || !r2) return new ExcelJsRange(this._ws, 1, 1, 1, 1);
            return new ExcelJsRange(
                this._ws,
                Math.min(r1.sr, r2.sr), Math.min(r1.sc, r2.sc),
                Math.max(r1.er, r2.er), Math.max(r1.ec, r2.ec),
            );
        }
        const p = parseRange(cell1);
        if (!p) return new ExcelJsRange(this._ws, 1, 1, 1, 1);
        return new ExcelJsRange(this._ws, p.sr, p.sc, p.er, p.ec);
    }

    get UsedRange(): ExcelJsRange {
        const rows = Math.max(this._ws.actualRowCount, 1);
        const cols = Math.max(this._ws.actualColumnCount, 1);
        return new ExcelJsRange(this._ws, 1, 1, rows, cols);
    }

    Rows(n: number | undefined = undefined): ExcelJsRange | { Count: number } {
        if (n !== undefined)
            return new ExcelJsRange(this._ws, n, 1, n, Math.max(this._ws.actualColumnCount, 256));
        // Rows.Count は VBA 仕様では 1048576 だが実用的な値を返す
        return { Count: 1048576 };
    }

    Columns(n: number | undefined = undefined): ExcelJsRange | { Count: number } {
        if (n !== undefined)
            return new ExcelJsRange(this._ws, 1, n, Math.max(this._ws.actualRowCount, 65536), n);
        return { Count: 16384 };
    }

    Select() {}
    Activate() {}

    get rawWorksheet(): ExcelJS.Worksheet { return this._ws; }
}

(ExcelJsWorksheet.prototype.Range as any).__vbaOverloads__ = [
    { params: [{ name: 'Cell1' }] },
    { params: [{ name: 'Cell1' }, { name: 'Cell2' }] },
] satisfies BuiltinOverload[];

// ---------- ExcelJsApplication ----------

export class ExcelJsApplication {
    private _sheets: Map<string, ExcelJsWorksheet> = new Map();
    private _activeSheetName: string;
    private _workbook: ExcelJS.Workbook;

    constructor(workbook: ExcelJS.Workbook) {
        this._workbook = workbook;
        workbook.eachSheet(ws => {
            this._sheets.set(ws.name, new ExcelJsWorksheet(ws));
        });
        const first = workbook.worksheets[0];
        this._activeSheetName = first?.name ?? 'Sheet1';
    }

    get ActiveSheet(): ExcelJsWorksheet { return this.Sheets(this._activeSheetName); }

    Sheets(nameOrIndex: string | number): ExcelJsWorksheet {
        if (typeof nameOrIndex === 'number') {
            const ws = this._workbook.getWorksheet(nameOrIndex);
            if (!ws) throw new Error(`Sheet index ${nameOrIndex} not found`);
            if (!this._sheets.has(ws.name))
                this._sheets.set(ws.name, new ExcelJsWorksheet(ws));
            return this._sheets.get(ws.name)!;
        }
        if (!this._sheets.has(nameOrIndex)) {
            const existing = this._workbook.getWorksheet(nameOrIndex);
            const ws = existing ?? this._workbook.addWorksheet(nameOrIndex);
            this._sheets.set(nameOrIndex, new ExcelJsWorksheet(ws));
        }
        return this._sheets.get(nameOrIndex)!;
    }

    Worksheets(nameOrIndex: string | number): ExcelJsWorksheet { return this.Sheets(nameOrIndex); }

    Range(cell1: string, cell2?: string): ExcelJsRange { return this.ActiveSheet.Range(cell1, cell2); }

    Cells(row: number, col: number): ExcelJsRange { return this.ActiveSheet.Cells(row, col); }

    Rows(...args: any[]) { return this.ActiveSheet.Rows(...args); }
    Columns(...args: any[]) { return this.ActiveSheet.Columns(...args); }

    setActiveSheet(name: string): void { this._activeSheetName = name; }

    get ActiveWorkbook() { return { Name: 'Workbook.xlsx', FullName: 'Workbook.xlsx', Path: '' }; }
    get ThisWorkbook() { return this.ActiveWorkbook; }

    get Workbooks() {
        const wb = this.ActiveWorkbook;
        return { Count: 1, Item: (_i: any) => wb, Open: (_p: string) => wb };
    }

    get Name() { return 'Microsoft Excel'; }
    get Version() { return '16.0'; }
    get Application(): this { return this; }

    get ScreenUpdating() { return true; }
    set ScreenUpdating(_v: any) {}
    get DisplayAlerts() { return true; }
    set DisplayAlerts(_v: any) {}
    get EnableEvents() { return true; }
    set EnableEvents(_v: any) {}
    get Calculation() { return -4105; }
    set Calculation(_v: any) {}
    get StatusBar() { return false as any; }
    set StatusBar(_v: any) {}
    get CutCopyMode() { return false as any; }
    set CutCopyMode(_v: any) {}
    get Visible() { return true; }
    set Visible(_v: any) {}

    get ActiveCell() { return this.ActiveSheet.Cells(1, 1); }
    get Selection() { return this.ActiveSheet.Cells(1, 1); }

    get WorksheetFunction(): Record<string, (...args: any[]) => any> {
        return new Proxy({} as Record<string, (...args: any[]) => any>, {
            get: (_t, prop) => {
                if (typeof prop === 'string') return (..._args: any[]) => 0;
            },
        });
    }
}

(ExcelJsApplication.prototype.Range as any).__vbaOverloads__ = [
    { params: [{ name: 'Cell1' }] },
    { params: [{ name: 'Cell1' }, { name: 'Cell2' }] },
] satisfies BuiltinOverload[];
