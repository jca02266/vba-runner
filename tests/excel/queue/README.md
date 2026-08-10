# Excel実機照合キュー

このディレクトリは、vba-runnerのVBA実行結果をWindows版Excelと照合するための
テストフィクスチャと実行スクリプトを置く場所である。vba-extractorの利用サンプルは
[`sample/excel/`](../../../sample/excel/)を参照する。

## ファイル

- `ExcelQueueVerification.bas`: XL-001〜XL-061の検証マクロ
- `ExcelQueueTicket.cls`: XL-033で使うクラスモジュール
- `ExcelQueueDefaultValue.cls`: XL-046で使う既定プロパティ付きクラス
- `ExcelQueueCallByNameTarget.cls`: XL-048で使うLet/Set片側プロパティ付きクラス
- `ExcelQueueForcedByVal.cls`、`ExcelQueueForcedByValInterface.cls`: XL-059のClass／Interface呼出し引数検証
- `empty_with_macro.xlsm`: マクロプロジェクトを持つ入力ブック
- `prepare-excel-vba.sh`: 開発側で`t.xlsm`を作成し、VBAをインポートするコマンド
- `eval-excel.cmd`: 準備済みブックをWindows Excelで実行するバッチ
- `run-excel-vba.ps1`: 指定したPublicプロシージャをExcelで実行する汎用ランナー
- `convert-to-utf8.ps1`: 結果ファイルをBOMなしUTF-8へ変換する汎用処理
- `finalize-excel-queue.ps1`: 完了マーカーを検証し、使用したVBAソースのハッシュを結果へ付加する処理
- `ExcelQueueVerification.result`: 最新の実機結果（UTF-8、BOMなし）

## ブックの準備と実機実行

`ExcelQueueVerification.bas`、`.cls`、`.frm`を更新したときは、まずNode環境のある
開発側（macOS/Linux等）で次を実行する。

```bash
tests/excel/queue/prepare-excel-vba.sh
```

このコマンドは`empty_with_macro.xlsm`を`t.xlsm`へコピーし、キューディレクトリ内の
VBAソース一式を`vba-extractor import`で`t.xlsm`へ取り込む。Windows側ではNodeや
vba-extractorを実行せず、準備済みの`t.xlsm`を使用する。

Windowsのコマンドプロンプトから、リポジトリ内のどこにいても次を実行できる。

```bat
tests\excel\queue\eval-excel.cmd
```

`eval-excel.cmd`は自身のディレクトリへ移動し、準備済み`t.xlsm`をExcelで開いて
`RunExcelQueueVerification`を実行する。開始時に
`%TEMP%\vba-runner-xl-queue`を削除するため、Binary Writeの短い出力が前回ファイルの
末尾データを引き継がない。

マクロは`ThisWorkbook.Path\ExcelQueueVerification.result`へ結果を直接書き込み、
`eval-excel.cmd`が終了後に`convert-to-utf8.ps1`でシステムコードページからBOMなしUTF-8へ変換する。
`Debug.Print`はImmediateウィンドウにも出力するが、記録の正本は結果ファイルである。

`prepare-excel-vba.sh`は、インポート対象の`.bas`、`.cls`、`.frm`をファイル名順に
並べ、改行をLFへ正規化したソース一式のSHA-256を計算し、その値をマクロへ埋め込んで
から`t.xlsm`を作成する。マクロは全プローブ終了後に埋め込み値を
`QUEUE_SOURCE_SHA256=<hash>`として出力し、その後`QUEUE_COMPLETE=True`を出力する。
`finalize-excel-queue.ps1`は結果に含まれるハッシュを現在のソースと照合するだけで、
実行後にハッシュを追記しない。これにより、古い`t.xlsm`の実行結果を現在のソースの
結果として受理できない。ソースを変更した場合は必ず準備スクリプトを再実行する。

## 個別実行

```powershell
powershell -ExecutionPolicy Bypass -File .\run-excel-vba.ps1 `
  -Workbook .\t.xlsm `
  -Module ExcelQueueVerification `
  -Procedure RunExcelQueueVerification
```

`-Module`は標準モジュール名、`-Procedure`はPublic Sub/Function名である。
ファイル名や`.bas`拡張子は指定しない。結果ファイルの生成・変換は呼び出し側の責務であり、
この汎用ランナーは出力パスを扱わない。

## 照合対象

- XL-001〜XL-010: 配列、UDT、文字列、TextStream
- XL-013〜XL-025: 比較、数値型、ファイルI/O、Lock、EOF
- XL-026〜XL-032: 非有限数値境界
- XL-033: クラス型配列の`As New`
- XL-034: Collection列挙中の`Add`とDictionary Keys列挙中の`Remove`
- XL-035〜XL-048: DateDiff、空配列、MIRR/SYD、On Error、Select Case、FormatNumber、TimeSerial、情報関数、過大リテラル、CallByNameの境界
- XL-059: Class／Interface member callのスペース付き括弧とByRef書戻し
- XL-061: Range.Value2由来の2次元Variant配列を動的Variant配列へ代入するshape境界

XL-023の逐次モードLock境界はExcelが待機する可能性があるため、一括実行では
スキップする。必要な場合だけ`RunExcelSequentialLockVerification`を単独実行し、
応答が戻らなければ未照合として扱う。
