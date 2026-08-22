# Excel実機照合キュー

このディレクトリは、vba-runnerのVBA実行結果をWindows版Excelと照合するための
テストフィクスチャと実行スクリプトを置く場所である。vba-extractorの利用サンプルは
[`sample/excel/`](../../../sample/excel)を参照する。

## ファイル

- `ExcelQueueVerification.bas`: XL-001〜XL-061の検証マクロ
- `FormatMatrixVerification.bas`: FormatのASCII単文字・複合トークン・字句境界マトリックス
- `RadixMatrixVerification.bas`: 変換型×基数×幅・符号・精度境界の直積マトリックス
- `ExcelQueueTicket.cls`: XL-033で使うクラスモジュール
- `ExcelQueueDefaultValue.cls`: XL-046で使う既定プロパティ付きクラス
- `ExcelQueueCallByNameTarget.cls`: XL-048で使うLet/Set片側プロパティ付きクラス
- `ExcelQueueForcedByVal.cls`、`ExcelQueueForcedByValInterface.cls`: XL-059のClass／Interface呼出し引数検証
- `empty_with_macro.xlsm`: マクロプロジェクトを持つ入力ブック
- `prepare-excel-vba.sh`: 開発側で`t.xlsm`を作成し、VBAをインポートするコマンド
- `verify-excel-queue-source.sh`: Windowsへ渡す前に準備スタンプを照合する検査
- `verify-excel-queue-source.ps1`: Windows側で準備スタンプと現在のソースを照合する事前検査
- `skip-if-excel-result-current.ps1`: 完了済み結果のグループハッシュを現在のソースと照合し、再実行要否を判定する処理
- `eval-excel.cmd`: 準備済みブックをWindows Excelで実行するバッチ
- `eval-matrix.cmd`: FormatとRadixのマトリックスを連続実行する専用バッチ
- `run-excel-vba.ps1`: 指定したPublicプロシージャをExcelで実行する汎用ランナー
- `convert-to-utf8.ps1`: 結果ファイルをBOMなしUTF-8へ変換する汎用処理
- `finalize-excel-queue.ps1`: 完了マーカーを検証し、使用したVBAソースのハッシュを結果へ付加する処理
- `ExcelQueueVerification.result`: 最新の実機結果（UTF-8、BOMなし）
- `FormatMatrix.result`: 専用Formatマトリックスの実機結果（通常回帰とは別管理）
- `RadixMatrix.result`: 変換境界マトリックスの実機結果（通常回帰とは別管理）

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

FormatとRadixの専用マトリックスも同じ準備済みブックで確認する場合は、通常キューとは
別に次を実行する。

```bat
tests\excel\queue\eval-matrix.cmd
```

このバッチはFormat、Radixの順に各Publicプロシージャを実行し、
`FormatMatrix.result`と`RadixMatrix.result`を個別に生成・検証する。

`eval-excel.cmd`は自身のディレクトリへ移動し、準備済み`t.xlsm`をExcelで開いて
`RunExcelQueueVerification`を実行する。開始時に
`%TEMP%\vba-runner-xl-queue`を削除するため、Binary Writeの短い出力が前回ファイルの
末尾データを引き継がない。

マクロは`ThisWorkbook.Path\ExcelQueueVerification.result`へ結果を直接書き込み、
`eval-excel.cmd`が終了後に`convert-to-utf8.ps1`でシステムコードページからBOMなしUTF-8へ変換する。
`Debug.Print`はImmediateウィンドウにも出力するが、記録の正本は結果ファイルである。

`prepare-excel-vba.sh`は、インポート対象の`.bas`、`.cls`、`.frm`をファイル名順に
並べ、改行をLFへ正規化したソース一式のSHA-256を計算し、その値を準備済みブックの
鮮度スタンプ（ハッシュ方式`grouped-v1`）として`t.xlsm.source.sha256`に保存する。さらに、結果ファイルに対応する
接頭辞（`ExcelQueue`、`FormatMatrix`、`RadixMatrix`）ごとにもSHA-256を計算し、各グループの
検証マクロへそのグループの値を埋め込む。マクロは全プローブ終了後に埋め込み値を
`QUEUE_SOURCE_SHA256=<hash>`として出力し、その後`QUEUE_COMPLETE=True`を出力する。
`finalize-excel-queue.ps1`と評価CLIは、結果名に対応する接頭辞のソースだけをハッシュして
結果に含まれる値と照合する。したがって、Formatだけを変更してもExcelQueueやRadixの
結果ハッシュは変わらず、変更していないテストの結果を更新する必要はない。
`finalize-excel-queue.ps1`は結果に含まれるハッシュを現在のソースと照合するだけで、
実行後にハッシュを追記しない。これにより、古い`t.xlsm`の実行結果を現在のソースの
結果として受理できない。`eval-excel.cmd`はExcelを起動する前にスタンプと現在の
ソースを照合するため、ソース変更後に準備を忘れた場合は明示的に停止する。ソースを
変更した場合は必ず準備スクリプトを再実行し、`t.xlsm`とスタンプをWindowsへ渡す。
さらに、既存の結果に完了マーカーがあり、結果の`QUEUE_SOURCE_SHA256`が現在の
グループソースと一致する場合は、Excelを起動せずその結果を再利用する。不一致、
結果欠落、または未完了の場合だけ結果を削除してExcelを実行する。
`eval-matrix.cmd`はFormatとRadixを個別に判定するため、一方だけソースが変わった場合も
もう一方の実機結果を再利用できる。

Formatの仕様探索は通常の回帰出力に混ぜない。`eval-matrix.cmd`は
`FormatMatrixVerification.bas`の`RunFormatMatrixVerification`と
`RadixMatrixVerification.bas`の`RunRadixMatrixVerification`を順に呼び出す。
前者はASCII印字可能文字
と代表的な数値・負数・ゼロ・Date・Currency・String引数の組み合わせ、複合トークン、および
実装の字句分岐から抽出した引用・エスケープ・角括弧・セクション・文字列プレースホルダー
境界を`FormatMatrix.result`へ出力し、後者は変換境界を`RadixMatrix.result`へ出力する。
Format仕様が固まった後の通常評価では、
この専用マトリックスを毎回実行しない。

Windowsへ渡す前には、開発側で次を実行する。

```bash
tests/excel/queue/verify-excel-queue-source.sh
```

この検査が成功した場合だけ、評価を`needs-excel-probe`から`needs-excel`へ遷移させる。

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
