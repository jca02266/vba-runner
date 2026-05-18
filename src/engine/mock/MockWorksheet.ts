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

interface CellRect {
    startCol: number;
    startRow: number;
    endCol: number;
    endRow: number;
}

export class MockRange {
    readonly __vbaDefault__ = true;

    constructor(private value: any) {}

    get Value(): any {
        return this.value;
    }

    set Value(val: any) {
        this.value = val;
    }

    valueOf(): any {
        return this.value;
    }
}

export class MockWorksheet {
    private name: string;
    private cells: Map<string, any> = new Map();

    constructor(name: string) {
        this.name = name;
    }

    /**
     * Range メソッド：Address で指定したセル/範囲を返す
     */
    Range(address: string): MockRange {
        const areas = this.resolveAddress(address);

        if (areas.length === 0) {
            return new MockRange(0);
        }

        if (areas.length === 1) {
            const area = areas[0];
            if (area.startCol === area.endCol && area.startRow === area.endRow) {
                // 単一セル — setter で cells を同期
                const cellAddr = this.cellAddress(area.startCol, area.startRow);
                const range = new MockRange(this.cells.get(cellAddr) ?? 0);
                const self = this;
                Object.defineProperty(range, 'Value', {
                    get: () => self.cells.get(cellAddr) ?? 0,
                    set: (val: any) => { self.cells.set(cellAddr, val); },
                });
                return range;
            }
            return new MockRange(this.getRectValues(area));
        }

        // 複数エリア（Union）: 全エリアの行を縦に結合して返す
        const combined: any[][] = [];
        for (const area of areas) {
            combined.push(...this.getRectValues(area));
        }
        return new MockRange(combined);
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
        return this.cells.get(address.toUpperCase().replace(/\$/g, '')) ?? 0;
    }

    /**
     * ワークシート内の全セルをダンプ（デバッグ用）
     */
    dump(): Record<string, any> {
        const result: Record<string, any> = {};
        this.cells.forEach((value, key) => { result[key] = value; });
        return result;
    }

    // ========== アドレス解析 ==========

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

    // ========== セル読み書きヘルパー ==========

    private getRectValues(area: CellRect): any[][] {
        const result: any[][] = [];
        for (let r = area.startRow; r <= area.endRow; r++) {
            const row: any[] = [];
            for (let c = area.startCol; c <= area.endCol; c++) {
                row.push(this.cells.get(this.cellAddress(c, r)) ?? 0);
            }
            result.push(row);
        }
        return result;
    }

    private setRectValue(area: CellRect, value: any): void {
        if (Array.isArray(value)) {
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
        } else {
            for (let r = area.startRow; r <= area.endRow; r++) {
                for (let c = area.startCol; c <= area.endCol; c++) {
                    this.cells.set(this.cellAddress(c, r), value);
                }
            }
        }
    }

    // ========== 基本ヘルパー ==========

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

    // ========== 基本プロパティ ==========

    get Name(): string {
        return this.name;
    }
}

export class MockApplication {
    private sheets: Map<string, MockWorksheet> = new Map();

    /**
     * Sheets('SheetName') でワークシートを取得
     * 存在しない場合は自動作成
     */
    Sheets(nameOrIndex: string | number): MockWorksheet {
        const name =
            typeof nameOrIndex === 'string' ? nameOrIndex : `Sheet${nameOrIndex}`;
        if (!this.sheets.has(name)) {
            this.sheets.set(name, new MockWorksheet(name));
        }
        return this.sheets.get(name)!;
    }

    listSheets(): string[] {
        return Array.from(this.sheets.keys());
    }

    clear(): void {
        this.sheets.clear();
    }
}
