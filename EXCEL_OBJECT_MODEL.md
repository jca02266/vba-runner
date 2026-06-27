# Excel オブジェクトモデル 実装状況一覧

## 目的

このドキュメントは、Excel VBA の標準オブジェクトモデルを網羅的に記録し、
以下3つの軸で実装状況を追跡・管理するためのロードマップである。

| 軸 | 対象ファイル |
|---|---|
| **exceljs 実現可能性** | exceljs ライブラリの機能として対応できるか |
| **アダプター実装** | `tools/exceljs-adapter/ExcelJsApplication.ts` での実装状況 |
| **モック実装** | `src/engine/mock/MockExcel.ts` / `MockWorksheet.ts` での実装状況 |

### 使い方

- 新規メンバーを実装したときは該当セルを ❌/⚠️ → ✅ に更新する
- 実装方針を決めた際は「備考」列に記載する
- スコープ外と決定した項目は ─ とし、理由を備考に書く

---

## 現状の制限

### 両実装共通

| 制限 | 詳細 |
|---|---|
| **WorksheetFunction が常に 0 を返す** | 数式評価エンジンを持たないため `WorksheetFunction.VLookup(...)` 等は全て 0。VBA 組み込み関数（`Len`, `UCase` 等）はエンジン実装済みなので影響なし |
| **イベントが発火しない** | `Worksheet_Change`, `Workbook_Open` 等のイベントプロシージャは呼ばれない |
| **名前付き範囲（Names）が未実装** | `Names.Add` / `Range("MyRange")` での名前参照は動作しない |
| **Chart / PivotTable / Shape は未実装** | これらを操作するコードは実行時エラーになる |
| **Application.InputBox / MsgBox は対話不可** | ヘッドレス環境のため実装しない方針。`MsgBox` は VBA 組み込み関数としてログ出力とデフォルト値返却のみ |

### モック（MockWorksheet / MockApplication）固有

| 制限 | 詳細 |
|---|---|
| **`End(xlUp/xlDown)` が常に Row=1 を返す** | セルデータを走査しないスタブのため。`ws.Cells(Rows.Count,1).End(xlUp).Row` で最終行を求めるパターンが正常動作しない |
| **`UsedRange` が A1 固定スタブ** | `ws.UsedRange.Rows.Count` で行数を得るパターンが正常動作しない |
| **`Range.Count` が常に 1** | 多セル範囲でも 1 を返す |
| **`Range.Address` が単一セル形式のみ正確** | 多セル範囲の `Address` は `$A$1` 固定（開始セルのみ） |
| **`Range.Offset` が getter（引数なし）のみ** | `Range.Offset(1, 0)` のように引数を渡しても正しく動作しない |
| **`Range.Resize` が引数なしスタブ** | 常に MockRange(0) を返す |
| **`Range.Clear/ClearContents` が no-op** | セルデータがクリアされない |
| **`Rows(n)` / `Columns(n)` で n 番目のデータを返さない** | 引数付きアクセスは MockRows/MockColumns スタブを返すだけで実際の行データを持たない |
| **`Worksheet.Index` が常に 1** | 複数シートでも全て Index=1 |
| **`Workbook` オブジェクトがスタブ** | Name/Path は固定文字列。Close/Save 等は未実装 |

### アダプター（ExcelJsApplication）固有

| 制限 | 詳細 |
|---|---|
| **色指定（`Font.Color`, `Interior.Color` 等）が no-op** | VBA の RGB 数値 → exceljs の ARGB 文字列変換が未実装 |
| **罫線（`Borders`）がスタブ** | `Range.Borders.Item(xlEdgeTop).LineStyle = xlContinuous` 等は no-op |
| **配置・折り返し（`HorizontalAlignment` 等）が no-op** | |
| **`Delete / Insert` が no-op** | 行・列シフトが発生しない。セルの値のみクリアされない |
| **`Copy / Paste / PasteSpecial` が no-op** | クリップボード経由のコピーは未実装 |
| **`Sort` が no-op** | セルデータのソートは行われない |
| **`Find` が常に Nothing を返す** | 検索機能は未実装 |
| **数式の再計算なし** | `cell.value` から取得できるのは保存済みキャッシュ値のみ。未計算セルは 0 |
| **`.xlsm` ラウンドトリップ不可** | exceljs は読み込み時に `vbaProject.bin` を破棄するため、マクロ付きブックとして保存できない |
| **`Worksheets.Count` 等コレクション操作が未実装** | `Worksheets.Count` / `For Each ws In Worksheets` は動作しない |
| **`Workbook` 経由の操作が未実装** | `ThisWorkbook.Worksheets("Sheet1")` のように Workbook を介したシートアクセスは動作しない（Application 直接アクセスは可） |

---

## 凡例

| 記号 | 意味 |
|---|---|
| ✅ | 実装済み（または exceljs で実現可能） |
| ⚠️ | スタブ／部分実装（メソッドは存在するが実際の操作はしない、またはデフォルト値固定） |
| ❌ | 未実装（メソッド・プロパティが存在しない） |
| ─ | 対象外（このプロジェクトのスコープ外と判断済み） |

| 種別記号 | 意味 |
|---|---|
| P | プロパティ (Property) |
| M | メソッド (Method) |
| C | コレクション (Collection / 子オブジェクトを返す) |

### 対象範囲の補足

Excel オブジェクトモデルは 500 以上のオブジェクト・数千のメンバーを持つ。
本ドキュメントは **VBA マクロで頻繁に使用されるオブジェクト・メンバー** を対象とする。
Chart・PivotTable・Shape 等の高度な機能は [スコープ外一覧](#スコープ外オブジェクト) に分類する。

---

## Application

VBA コードからは `Application.xxx` または修飾なし（`Range("A1")` 等）でアクセス。
エンジンでは `setDefaultBindingObject(app)` と `setBuiltinOverride('Application', app)` で注入。

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `ActiveSheet` | P | ─ | ✅ | ✅ | |
| `ActiveWorkbook` | P | ─ | ⚠️ | ⚠️ | 固定スタブ（Name/Path/FullName のみ） |
| `ThisWorkbook` | P | ─ | ⚠️ | ⚠️ | 固定スタブ |
| `ActiveCell` | P | ─ | ⚠️ | ⚠️ | 常に A1 を返す |
| `Selection` | P | ─ | ⚠️ | ⚠️ | 常に A1 を返す |
| `Name` | P | ─ | ✅ | ✅ | `"Microsoft Excel"` 固定 |
| `Version` | P | ─ | ✅ | ✅ | `"16.0"` 固定 |
| `Application` | P | ─ | ✅ | ✅ | 自己参照 |
| `ScreenUpdating` | P | ─ | ⚠️ | ⚠️ | 読み書きともにno-op |
| `DisplayAlerts` | P | ─ | ⚠️ | ⚠️ | no-op |
| `EnableEvents` | P | ─ | ⚠️ | ⚠️ | no-op |
| `Calculation` | P | ─ | ⚠️ | ⚠️ | no-op（定数 -4105 を返す） |
| `StatusBar` | P | ─ | ⚠️ | ⚠️ | no-op |
| `CutCopyMode` | P | ─ | ⚠️ | ⚠️ | no-op |
| `Visible` | P | ─ | ⚠️ | ⚠️ | no-op |
| `OperatingSystem` | P | ─ | ❌ | ❌ | |
| `Path` | P | ─ | ❌ | ❌ | |
| `Caption` | P | ─ | ❌ | ❌ | |
| `UserName` | P | ─ | ❌ | ❌ | |
| `WorksheetFunction` | C | ─ | ⚠️ | ⚠️ | Proxy で常に 0 を返す（[詳細](#worksheetfunction)） |
| `Workbooks` | C | ─ | ⚠️ | ⚠️ | Count=1 固定、Open() はスタブ |
| `Sheets()` | M | ─ | ✅ | ✅ | 名前・インデックス両対応 |
| `Worksheets()` | M | ─ | ✅ | ✅ | Sheets() の別名 |
| `Range()` | M | ─ | ✅ | ✅ | ActiveSheet に委譲 |
| `Cells()` | M | ─ | ✅ | ✅ | ActiveSheet に委譲 |
| `Rows()` | M | ─ | ✅ | ✅ | ActiveSheet に委譲 |
| `Columns()` | M | ─ | ✅ | ✅ | ActiveSheet に委譲 |
| `InputBox()` | M | ─ | ❌ | ❌ | 対話不可環境のため実装しない方針 |
| `OnKey()` | M | ─ | ❌ | ❌ | |
| `OnTime()` | M | ─ | ❌ | ❌ | |
| `GoTo()` | M | ─ | ❌ | ❌ | |
| `Calculate()` | M | ─ | ❌ | ❌ | |
| `Quit()` | M | ─ | ❌ | ❌ | |
| `Wait()` | M | ─ | ❌ | ❌ | |
| `Run()` | M | ─ | ─ | ─ | VBARunner.run() で代替 |

---

## Workbooks コレクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Count` | P | ─ | ⚠️ | ⚠️ | 常に 1 |
| `Item()` | M | ─ | ⚠️ | ⚠️ | 常に同一スタブを返す |
| `Open()` | M | ✅ | ⚠️ | ⚠️ | スタブ（ファイルを開かない） |
| `Add()` | M | ✅ | ❌ | ❌ | |
| `Close()` | M | ─ | ❌ | ❌ | |
| `For Each` 反復 | ─ | ─ | ❌ | ❌ | |

---

## Workbook

`ThisWorkbook` / `ActiveWorkbook` 経由でアクセス。現状はスタブオブジェクトのみ。

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Name` | P | ✅ | ⚠️ | ⚠️ | 固定文字列 |
| `Path` | P | ✅ | ⚠️ | ⚠️ | 空文字固定 |
| `FullName` | P | ✅ | ⚠️ | ⚠️ | Name と同一 |
| `Saved` | P | ─ | ❌ | ❌ | |
| `ReadOnly` | P | ─ | ❌ | ❌ | |
| `Sheets` | C | ✅ | ❌ | ❌ | Application 経由でのみアクセス可 |
| `Worksheets` | C | ✅ | ❌ | ❌ | 同上 |
| `ActiveSheet` | P | ✅ | ❌ | ❌ | Application.ActiveSheet で代替 |
| `Names` | C | ✅ | ❌ | ❌ | |
| `Close()` | M | ✅ | ❌ | ❌ | |
| `Save()` | M | ✅ | ❌ | ❌ | 呼び出し元で wb.xlsx.writeFile() |
| `SaveAs()` | M | ✅ | ❌ | ❌ | |
| `Activate()` | M | ─ | ❌ | ❌ | |
| `Protect()` | M | ✅ | ❌ | ❌ | |
| `Unprotect()` | M | ✅ | ❌ | ❌ | |

---

## Worksheets / Sheets コレクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Count` | P | ✅ | ❌ | ❌ | コレクションオブジェクト未実装。Sheets() メソッドで代替 |
| `Item()` | M | ✅ | ✅ | ✅ | `Sheets(name\|index)` で実現 |
| `Add()` | M | ✅ | ⚠️ | ⚠️ | 存在しないシート名でアクセスすると自動作成 |
| `Copy()` | M | ✅ | ❌ | ❌ | |
| `Delete()` | M | ✅ | ❌ | ❌ | |
| `For Each` 反復 | ─ | ✅ | ❌ | ❌ | |

---

## Worksheet

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Name` | P | ✅ | ✅ | ✅ | 読み取り専用（書き込みは未実装） |
| `Index` | P | ✅ | ❌ | ⚠️ | モックは常に 1 |
| `CodeName` | P | ─ | ❌ | ❌ | |
| `UsedRange` | P | ✅ | ✅ | ⚠️ | モックは A1 固定スタブ |
| `Visible` | P | ✅ | ❌ | ⚠️ | モックは no-op |
| `ProtectContents` | P | ✅ | ❌ | ❌ | |
| `EnableSelection` | P | ─ | ❌ | ⚠️ | モックは no-op |
| `Cells()` | M | ✅ | ✅ | ✅ | |
| `Range()` | M | ✅ | ✅ | ✅ | 2引数形式・Union・Intersection 対応（モックのみ） |
| `Rows()` | C | ✅ | ✅ | ✅ | |
| `Columns()` | C | ✅ | ✅ | ✅ | |
| `Names` | C | ✅ | ❌ | ❌ | |
| `Comments` | C | ✅ | ❌ | ❌ | |
| `Hyperlinks` | C | ✅ | ❌ | ❌ | |
| `ListObjects` | C | ✅ | ❌ | ❌ | Excel テーブル（Table機能） |
| `AutoFilter` | P/M | ✅ | ❌ | ❌ | |
| `Sort` | C | ✅ | ❌ | ❌ | Sort オブジェクト（Range.Sort とは別） |
| `Activate()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `Select()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `Copy()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Move()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Delete()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Calculate()` | M | ─ | ❌ | ❌ | |
| `Protect()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Unprotect()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `ShowAllData()` | M | ✅ | ❌ | ❌ | |
| `SetBackgroundPicture()` | M | ─ | ─ | ─ | |

---

## Range

最も重要なオブジェクト。

### 値・数式

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Value` | P | ✅ | ✅ | ✅ | スカラー・2D 配列一括代入に対応 |
| `Value2` | P | ✅ | ❌ | ❌ | Value と同等（日付を数値として返す点のみ差異） |
| `Formula` | P | ✅ | ✅ | ❌ | アダプターは `=A1+B1` 形式で書き込み |
| `FormulaR1C1` | P | ⚠️ | ❌ | ❌ | exceljs は R1C1 形式を直接サポートしない |
| `FormulaLocal` | P | ─ | ❌ | ❌ | |
| `Text` | P | ⚠️ | ❌ | ❌ | 書式適用後の文字列表現 |
| `HasFormula` | P | ✅ | ❌ | ❌ | |

### アドレス・位置

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Address` | P | ✅ | ✅ | ⚠️ | モックは単一セル形式のみ正確（`$A$1`） |
| `AddressLocal` | P | ─ | ❌ | ❌ | |
| `Row` | P | ✅ | ✅ | ✅ | |
| `Column` | P | ✅ | ✅ | ✅ | |
| `Count` | P | ✅ | ✅ | ⚠️ | モックは常に 1 |
| `Rows` | C | ✅ | ✅ | ✅ | `.Count` 取得・`(n)` で行 Range を返す |
| `Columns` | C | ✅ | ✅ | ✅ | `.Count` 取得・`(n)` で列 Range を返す |
| `EntireRow` | P | ✅ | ✅ | ⚠️ | モックはスタブ Range |
| `EntireColumn` | P | ✅ | ✅ | ⚠️ | モックはスタブ Range |
| `CurrentRegion` | P | ❌ | ❌ | ❌ | データ周囲の連続セル範囲 |
| `Offset()` | M | ✅ | ✅ | ⚠️ | モックは getter のみ（引数なし） |
| `Resize()` | M | ✅ | ✅ | ⚠️ | モックは引数なしのスタブ |
| `End()` | M | ─ | ✅ | ⚠️ | アダプターは xlUp/xlDown 実装。モックは常に Row=1 |
| `Areas` | C | ─ | ❌ | ❌ | Union 範囲のサブエリアコレクション |
| `MergeArea` | P | ✅ | ❌ | ❌ | |
| `Next` | P | ─ | ❌ | ❌ | |
| `Previous` | P | ─ | ❌ | ❌ | |
| `Cells()` | M | ✅ | ❌ | ❌ | Range 内の相対 Cells アクセス |
| `Parent` | P | ─ | ❌ | ❌ | 親 Worksheet への参照 |

### 書式

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `NumberFormat` | P | ✅ | ✅ | ⚠️ | モックは no-op（`"General"` 固定） |
| `NumberFormatLocal` | P | ✅ | ✅ | ⚠️ | モックは no-op |
| `HorizontalAlignment` | P | ✅ | ⚠️ | ⚠️ | 両者 no-op |
| `VerticalAlignment` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `WrapText` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `ShrinkToFit` | P | ✅ | ❌ | ❌ | |
| `Orientation` | P | ⚠️ | ❌ | ❌ | exceljs は alignment.textRotation |
| `IndentLevel` | P | ✅ | ❌ | ❌ | |
| `MergeCells` | P | ✅ | ⚠️ | ⚠️ | no-op（Merge/UnMerge との組み合わせは未実装） |
| `Hidden` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `Locked` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `RowHeight` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `ColumnWidth` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `Font` | C | ✅ | ✅ | ⚠️ | [Font 詳細](#font) 参照 |
| `Interior` | C | ✅ | ⚠️ | ⚠️ | [Interior 詳細](#interior) 参照 |
| `Borders` | C | ✅ | ⚠️ | ⚠️ | [Borders 詳細](#borders--border) 参照 |
| `Style` | P | ✅ | ❌ | ❌ | セルスタイル名 |
| `FormatConditions` | C | ✅ | ❌ | ❌ | 条件付き書式コレクション |
| `Validation` | C | ✅ | ❌ | ❌ | 入力規則 |

### アクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Select()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `Activate()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `Copy()` | M | ✅ | ⚠️ | ⚠️ | no-op（コピー操作は未実装） |
| `Cut()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `Paste()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `PasteSpecial()` | M | ⚠️ | ⚠️ | ⚠️ | no-op |
| `Delete()` | M | ✅ | ⚠️ | ⚠️ | no-op（行・列削除・シフトなし） |
| `Insert()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `Clear()` | M | ✅ | ✅ | ⚠️ | アダプターは値クリア実装。モックは no-op |
| `ClearContents()` | M | ✅ | ✅ | ⚠️ | アダプターは値クリア実装。モックは no-op |
| `ClearFormats()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `ClearComments()` | M | ✅ | ❌ | ❌ | |
| `Merge()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `UnMerge()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `AutoFit()` | M | ✅ | ⚠️ | ⚠️ | no-op |
| `AutoFill()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `Sort()` | M | ✅ | ⚠️ | ⚠️ | no-op（実際のソートなし） |
| `FillDown()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `FillRight()` | M | ─ | ⚠️ | ⚠️ | no-op |
| `FillUp()` | M | ─ | ❌ | ❌ | |
| `FillLeft()` | M | ─ | ❌ | ❌ | |
| `Find()` | M | ─ | ⚠️ | ⚠️ | 常に null を返す |
| `FindNext()` | M | ─ | ❌ | ❌ | |
| `Replace()` | M | ─ | ❌ | ❌ | |
| `AutoFilter()` | M | ✅ | ❌ | ❌ | |
| `Group()` | M | ✅ | ❌ | ❌ | |
| `Ungroup()` | M | ✅ | ❌ | ❌ | |
| `AddComment()` | M | ✅ | ❌ | ❌ | |
| `SpecialCells()` | M | ─ | ❌ | ❌ | |
| `TextToColumns()` | M | ─ | ❌ | ❌ | |
| `GoalSeek()` | M | ─ | ─ | ─ | |
| `Calculate()` | M | ─ | ❌ | ❌ | |
| `DataSeries()` | M | ─ | ❌ | ❌ | |

---

## Font

`Range.Font` 経由でアクセス。

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Bold` | P | ✅ | ✅ | ⚠️ | モックは no-op |
| `Italic` | P | ✅ | ✅ | ⚠️ | モックは no-op |
| `Size` | P | ✅ | ✅ | ⚠️ | モックは no-op |
| `Name` | P | ✅ | ✅ | ⚠️ | フォント名。モックは no-op |
| `Underline` | P | ✅ | ⚠️ | ⚠️ | exceljs は `single/double/...`。両者 no-op |
| `Color` | P | ✅ | ⚠️ | ⚠️ | RGB 数値。exceljs は ARGB 文字列への変換要 |
| `ColorIndex` | P | ⚠️ | ⚠️ | ⚠️ | カラーパレットインデックス。no-op |
| `ThemeColor` | P | ✅ | ❌ | ❌ | |
| `TintAndShade` | P | ✅ | ❌ | ❌ | |
| `StrikeThrough` | P | ✅ | ❌ | ❌ | exceljs では `strike` |
| `Superscript` | P | ✅ | ❌ | ❌ | exceljs では `vertAlign: "superscript"` |
| `Subscript` | P | ✅ | ❌ | ❌ | exceljs では `vertAlign: "subscript"` |
| `Background` | P | ─ | ❌ | ❌ | |

---

## Interior

`Range.Interior` 経由でアクセス。VBA でのセル背景色設定。

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Color` | P | ✅ | ⚠️ | ⚠️ | exceljs は `cell.fill`（fgColor）。変換実装が必要。no-op |
| `ColorIndex` | P | ⚠️ | ⚠️ | ⚠️ | no-op |
| `Pattern` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `PatternColor` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `PatternColorIndex` | P | ⚠️ | ❌ | ❌ | |
| `ThemeColor` | P | ✅ | ❌ | ❌ | |
| `TintAndShade` | P | ✅ | ❌ | ❌ | |

---

## Borders / Border

`Range.Borders` 経由でアクセス。

### Borders コレクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Item(index)` | M | ✅ | ⚠️ | ✅ | exceljs は `cell.border`。定数は xlEdgeLeft(-4131)等 |
| `LineStyle` | P | ✅ | ⚠️ | ⚠️ | コレクション全体への一括設定。no-op |
| `Weight` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `Color` | P | ✅ | ⚠️ | ⚠️ | no-op |

### Border（単一辺）

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `LineStyle` | P | ✅ | ⚠️ | ⚠️ | xlContinuous(1)/xlDash(-4115)等 → exceljs 形式に変換要 |
| `Weight` | P | ✅ | ⚠️ | ⚠️ | xlThin(2)/xlMedium(-4138)/xlThick(4)等 |
| `Color` | P | ✅ | ⚠️ | ⚠️ | no-op |
| `ColorIndex` | P | ⚠️ | ⚠️ | ⚠️ | no-op |

---

## Rows / Columns コレクション

`Worksheet.Rows` / `Range.Rows` 等で取得。

### Rows コレクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Count` | P | ✅ | ✅ | ✅ | 1048576 固定 |
| `(n)` アクセス | ─ | ✅ | ✅ | ⚠️ | アダプターは実 Range。モックは MockRows スタブ |
| `Hidden` | P | ✅ | ❌ | ⚠️ | モックは no-op |
| `RowHeight` | P | ✅ | ❌ | ⚠️ | モックは no-op |
| `AutoFit()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Select()` | M | ─ | ❌ | ⚠️ | モックは no-op |
| `Delete()` | M | ✅ | ❌ | ⚠️ | モックは no-op |

### Columns コレクション

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Count` | P | ✅ | ✅ | ✅ | 16384 固定 |
| `(n)` アクセス | ─ | ✅ | ✅ | ⚠️ | アダプターは実 Range。モックは MockColumns スタブ |
| `Hidden` | P | ✅ | ❌ | ⚠️ | モックは no-op |
| `ColumnWidth` | P | ✅ | ❌ | ⚠️ | モックは no-op |
| `AutoFit()` | M | ✅ | ❌ | ⚠️ | モックは no-op |
| `Select()` | M | ─ | ❌ | ⚠️ | モックは no-op |
| `Delete()` | M | ✅ | ❌ | ⚠️ | モックは no-op |

---

## WorksheetFunction

`WorksheetFunction.Sum(...)` 等でアクセス。

exceljs は数式の計算エンジンを持たない。計算が必要な場合は hyperformula 等の別ライブラリが必要。
現状は両実装とも Proxy で全メソッドに対して `0` を返す。

| カテゴリ | 代表的な関数 | exceljs | アダプター | モック |
|---|---|---|---|---|
| 数学・集計 | Sum, Average, Min, Max, Count, CountA, CountBlank, Round, Abs, Int, Mod, Sqrt | ❌ | ⚠️ | ⚠️ |
| 条件集計 | CountIf, CountIfs, SumIf, SumIfs, AverageIf, AverageIfs | ❌ | ⚠️ | ⚠️ |
| 検索・参照 | VLookup, HLookup, Index, Match, Offset, Choose | ❌ | ⚠️ | ⚠️ |
| 文字列 | Len, Left, Right, Mid, Trim, Upper, Lower, Substitute, Find, Search, Concatenate | ❌ | ⚠️ | ⚠️ |
| 日付・時刻 | Now, Today, Year, Month, Day, DateDiff, EoMonth, WorkDay, NetworkDays | ❌ | ⚠️ | ⚠️ |
| 論理 | And, Or, Not, IfError, IsNA, IsNumber, IsText, IsEmpty | ❌ | ⚠️ | ⚠️ |
| 統計 | StDev, Var, Median, Mode, Rank, Large, Small, Percentile, Correl | ❌ | ⚠️ | ⚠️ |
| 財務 | PV, FV, NPV, IRR, PMT | ❌ | ⚠️ | ⚠️ |

> **注**: VBA 組み込み関数（`Len`, `Left`, `UCase` 等）はエンジンが直接実装しているため対象外。
> WorksheetFunction は `Application.WorksheetFunction.VLookup(...)` 形式でのみ使用するもの。

---

## Comment / Comments

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Range.AddComment()` | M | ✅ | ❌ | ❌ | exceljs は `cell.note` / `addNote()` |
| `Comment.Text()` | M | ✅ | ❌ | ❌ | |
| `Comment.Delete()` | M | ✅ | ❌ | ❌ | |
| `Comment.Visible` | P | ✅ | ❌ | ❌ | |
| `Comments.Count` | P | ✅ | ❌ | ❌ | |

---

## Name / Names（名前付き範囲）

| メンバー | 種別 | exceljs | アダプター | モック | 備考 |
|---|---|---|---|---|---|
| `Names.Add()` | M | ✅ | ❌ | ❌ | exceljs は `workbook.definedNames.add()` |
| `Names.Item()` | M | ✅ | ❌ | ❌ | |
| `Names.Count` | P | ✅ | ❌ | ❌ | |
| `Name.Name` | P | ✅ | ❌ | ❌ | |
| `Name.RefersTo` | P | ✅ | ❌ | ❌ | |
| `Name.Delete()` | M | ✅ | ❌ | ❌ | |

---

## スコープ外オブジェクト

以下はリファクタリング・ユニットテスト支援という本プロジェクトのスコープ外と判断。
exceljs での実現可能性のみ記録する。

| オブジェクト | exceljs | 備考 |
|---|---|---|
| `Chart` / `Charts` | ⚠️ | exceljs は基本的なチャートの追加のみ。既存チャートの読み取り・操作は限定的 |
| `PivotTable` / `PivotTables` | ❌ | exceljs は PivotTable を読み取り専用で保持するのみ |
| `Shape` / `Shapes` | ⚠️ | 画像の挿入のみ対応。図形・テキストボックスは不可 |
| `ListObject` / `ListObjects` | ✅ | exceljs はテーブルの読み書きに対応 |
| `AutoFilter` / `Filter` | ✅ | exceljs は AutoFilter の設定を保持 |
| `FormatCondition` / `FormatConditions` | ✅ | exceljs は条件付き書式の読み書きに対応 |
| `Hyperlink` / `Hyperlinks` | ✅ | exceljs は `cell.value` に URL リンクを設定可能 |
| `Validation` (入力規則) | ✅ | exceljs は `cell.dataValidation` に対応 |
| `QueryTable` | ❌ | |
| `OLEObject` / `OLEObjects` | ❌ | |
| `Window` / `Windows` | ─ | 画面操作。ヘッドレス環境では不要 |
| `UserForm` / `Controls` | ─ | GUI 操作。ヘッドレス環境では不要 |
| `CommandBar` / `Ribbon` | ─ | UI カスタマイズ。スコープ外 |

---

*最終更新: 2026-06-28*
