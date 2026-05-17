/**
 * 軽量な Excel Worksheet モック
 *
 * 使い方:
 * const mockWs = new MockWorksheet('Sheet1');
 * mockWs.setCellValue('A1', 100);
 * mockWs.setCellValue('B1:B5', [[1], [2], [3], [4], [5]]);
 */

export class MockRange {
  constructor(private value: any) {}

  get Value(): any {
    return this.value;
  }

  set Value(val: any) {
    this.value = val;
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
   * 対応フォーマット:
   * - 'A1' （単一セル）
   * - 'A1:B5' （範囲：配列として返す）
   */
  Range(address: string): MockRange {
    const normalized = address.toUpperCase();

    // 範囲指定（A1:B5）
    if (normalized.includes(':')) {
      return new MockRange(this.getRangeValues(normalized));
    }

    // 単一セル
    const value = this.cells.get(normalized) ?? 0;
    const range = new MockRange(value);

    // セッタをオーバーライド：値変更時に cells を更新
    const self = this;
    Object.defineProperty(range, 'Value', {
      get: () => self.cells.get(normalized) ?? 0,
      set: (val: any) => {
        self.cells.set(normalized, val);
      },
    });

    return range;
  }

  /**
   * セルに値を設定
   * @param address セルアドレス ('A1') または範囲 ('A1:A5')
   * @param value スカラー値または配列
   */
  setCellValue(address: string, value: any): void {
    const normalized = address.toUpperCase();

    if (normalized.includes(':')) {
      // 範囲設定
      const [start, end] = normalized.split(':');
      const [startCol, startRow] = this.parseAddress(start);
      const [endCol, endRow] = this.parseAddress(end);

      if (Array.isArray(value)) {
        // 配列で複数セルを設定
        let arrayRow = 0;
        for (let r = startRow; r <= endRow; r++) {
          const rowData = value[arrayRow];
          if (Array.isArray(rowData)) {
            let arrayCol = 0;
            for (let c = startCol; c <= endCol; c++) {
              const cellAddr = this.cellAddress(c, r);
              this.cells.set(cellAddr, rowData[arrayCol] ?? '');
              arrayCol++;
            }
          }
          arrayRow++;
        }
      } else {
        // スカラー値で範囲全体を設定
        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const cellAddr = this.cellAddress(c, r);
            this.cells.set(cellAddr, value);
          }
        }
      }
    } else {
      // 単一セル設定
      this.cells.set(normalized, value);
    }
  }

  /**
   * セルから値を取得
   */
  getCellValue(address: string): any {
    return this.cells.get(address.toUpperCase()) ?? 0;
  }

  /**
   * ワークシート内の全セルをダンプ（デバッグ用）
   */
  dump(): Record<string, any> {
    const result: Record<string, any> = {};
    this.cells.forEach((value, key) => {
      result[key] = value;
    });
    return result;
  }

  // ========== ヘルパーメソッド ==========

  private getRangeValues(range: string): any[][] {
    const [start, end] = range.split(':');
    const [startCol, startRow] = this.parseAddress(start);
    const [endCol, endRow] = this.parseAddress(end);

    const result: any[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: any[] = [];
      for (let c = startCol; c <= endCol; c++) {
        const cellAddr = this.cellAddress(c, r);
        row.push(this.cells.get(cellAddr) ?? 0);
      }
      result.push(row);
    }
    return result;
  }

  private parseAddress(address: string): [number, number] {
    // 'A1' → [1, 1], 'B5' → [2, 5]
    const match = address.match(/([A-Z]+)(\d+)/);
    if (!match) return [1, 1];

    const col = this.columnToNumber(match[1]);
    const row = parseInt(match[2], 10);
    return [col, row];
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

  /**
   * ワークシート一覧を取得（デバッグ用）
   */
  listSheets(): string[] {
    return Array.from(this.sheets.keys());
  }

  /**
   * すべてを初期化
   */
  clear(): void {
    this.sheets.clear();
  }
}
