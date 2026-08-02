# Excel実機照合キュー

このディレクトリは、vba-runnerのVBA実行結果をWindows版Excelと照合するための
テストフィクスチャと実行スクリプトを置く場所である。vba-extractorの利用サンプルは
[`sample/excel/`](../../../sample/excel/)を参照する。

## ファイル

- `ExcelQueueVerification.bas`: XL-001〜XL-034の検証マクロ
- `ExcelQueueTicket.cls`: XL-033で使うクラスモジュール
- `empty_with_macro.xlsm`: マクロプロジェクトを持つ入力ブック
- `eval-excel.cmd`: ブック複製、VBAインポート、Excel実行を一括するWindowsバッチ
- `run-excel-queue.ps1`: 指定したPublicプロシージャをExcelで実行する汎用ランナー
- `ExcelQueueVerification.result`: 最新の実機結果（UTF-8、BOMなし）

## 一括実行

Windowsのコマンドプロンプトから、リポジトリ内のどこにいても次を実行できる。

```bat
tests\excel\queue\eval-excel.cmd
```

バッチは自身のディレクトリへ移動し、`empty_with_macro.xlsm`を`t.xlsm`へ複製する。
その後、`ExcelQueueVerification.bas`と`ExcelQueueTicket.cls`をインポートして
`RunExcelQueueVerification`を実行する。開始時に
`%TEMP%\vba-runner-xl-queue`を削除するため、Binary Writeの短い出力が前回ファイルの
末尾データを引き継がない。

マクロは`ThisWorkbook.Path\ExcelQueueVerification.result`へ結果を直接書き込み、
PowerShellが終了後にシステムコードページからBOMなしUTF-8へ変換する。
`Debug.Print`はImmediateウィンドウにも出力するが、記録の正本は結果ファイルである。

## 個別実行

```powershell
powershell -ExecutionPolicy Bypass -File .\run-excel-queue.ps1 `
  -Workbook .\t.xlsm `
  -Module ExcelQueueVerification `
  -Procedure RunExcelQueueVerification
```

`-Module`は標準モジュール名、`-Procedure`はPublic Sub/Function名である。
ファイル名や`.bas`拡張子は指定しない。`-Output`を省略すると、ブックと同じディレクトリの
`ExcelQueueVerification.result`を確認する。

## 照合対象

- XL-001〜XL-010: 配列、UDT、文字列、TextStream
- XL-013〜XL-025: 比較、数値型、ファイルI/O、Lock、EOF
- XL-026〜XL-032: 非有限数値境界
- XL-033: クラス型配列の`As New`
- XL-034: Collection列挙中の`Add`とDictionary Keys列挙中の`Remove`

XL-023の逐次モードLock境界はExcelが待機する可能性があるため、一括実行では
スキップする。必要な場合だけ`RunExcelSequentialLockVerification`を単独実行し、
応答が戻らなければ未照合として扱う。
