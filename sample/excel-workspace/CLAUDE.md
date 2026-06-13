# excel-workspace CLAUDE.md

vba-runner の評価を行う Excel VBA シミュレーションワークスペースです。

**重要**: このワークスペース以外の情報を見ないでください
実行、テストにあたって *.bas や *.cls などの VBA ソースを修正しないでください

作業報告、追加情報、バグ報告を REPORT.md に記録してください

## ワークスペース構成

| ファイル | 役割 |
|---|---|
| `SimCell.cls` | Excel のセル (Range 単一セル) シミュレーション |
| `SimSheet.cls` | Excel の Worksheet シミュレーション (Change/Activate イベント) |
| `SimWorkbook.cls` | Excel の Workbook シミュレーション (Open/BeforeClose/SheetChange イベント) |
| `ThisWorkbook.cls` | Excel の `ThisWorkbook` モジュール模倣 (Workbook_Open 等) |
| `Sheet1.cls` | Excel の `Sheet1` モジュール模倣 (Worksheet_Change 等) |
| `ExcelUtils.bas` | AutoSort / VLookup 相当 / SumIf 等の標準モジュール |
| `ExcelMain.bas` | デモ実行エントリーポイント |

## デモ手順

- `RunWorkbookDemo` — Workbook_Open → データ入力 → SheetChange → BeforeClose のイベントチェーン
- `RunDataDemo`     — ソート・集計・VLookup・エラーハンドリング・レポート生成
- `RunValidationDemo` — Worksheet_Change による入力検証 (空白・負数・文字列チェック)
