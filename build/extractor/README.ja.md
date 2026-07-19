# vba-extractor

Office ファイル（.xlsm など）から VBA ソースを抽出・書き戻しする CLI ツールです。
Excel を使わずにコマンドラインで VBA ソースをテキストファイルとして管理できます。

## インストール

```bash
npm install -g vba-extractor
```

## 使い方

### export — VBA ソースをファイルに抽出

```bash
vba-extractor export <input.xlsm> [output-dir] [--encoding <cp>]
```

| 引数 | 説明 |
|---|---|
| `<input.xlsm>` | 対象の Office ファイル |
| `[output-dir]` | 出力先ディレクトリ（省略時は input と同じディレクトリ） |
| `--encoding <cp>` | コードページを明示指定（例: `cp932`, `cp1252`）。省略時はファイル内の PROJECTCODEPAGE を使用 |

```bash
# カレントディレクトリに .bas / .cls を出力
vba-extractor export MyBook.xlsm

# 出力先を指定
vba-extractor export MyBook.xlsm src/vba
```

### import — VBA ソースを Office ファイルに書き戻し

```bash
vba-extractor import <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>] [--yes]
```

| 引数 | 説明 |
|---|---|
| `<input.xlsm>` | 対象の Office ファイル |
| `<source-dir>` | 書き戻す .bas / .cls を格納したディレクトリ |
| `[output.xlsm]` | 出力ファイル（省略時は input を上書き） |
| `--encoding <cp>` | コードページを明示指定。省略時はファイル内の PROJECTCODEPAGE を使用 |
| `--yes`, `-y` | 確認プロンプト（`Create a backup and run import? [y/N]`）をスキップ。スクリプト・AIエージェントから繰り返し実行する場合に使う |

```bash
# 編集した .bas / .cls を xlsm に書き戻す
vba-extractor import MyBook.xlsm src/vba

# 元ファイルを残して別ファイルに出力
vba-extractor import MyBook.xlsm src/vba MyBook_updated.xlsm

# 確認プロンプト無しで実行（自動化スクリプト向け）
vba-extractor import MyBook.xlsm src/vba --yes
```

VBE の「ファイルのエクスポート」メニューで出力した `.cls` ファイルには `VERSION 1.0 CLASS` /
`BEGIN ... END`（`MultiUse` 等のコンポーネント属性を保持するブロック）というヘッダーが付いています。
これは `export` コマンドの出力には含まれず、モジュールソース自体にも本来含まれるべきではありません。
`import` はこのヘッダーを自動検出して取り除くため、どちらの形式の `.cls` ファイルも手動編集なしで
そのまま入力として使えます。

### 新規 .xlsm をゼロから作る (Windows + Excel)

`import` は `xl/vbaProject.bin` を含む既存の `.xlsm` が前提で、新規に `.xlsm` を作ることはできません。マクロ有効形式で保存しただけで VBA プロジェクトを一度も作成していない `.xlsm` には、このパートが存在しません。`.bas` ファイルしか
無い状態(AIが生成した、または手書きで、まだExcelワークブックが無い場合)から始めるときは、
同梱のPowerShellスクリプトでExcelのCOM自動化を使って最初の `.xlsm` を作成し、その後
`import` でクラスモジュール等の残りを同期します:

```powershell
# Windows + ローカルにインストールされたExcel、かつExcelの
# 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」(AccessVBOM) 設定が有効である必要があります。
powershell -File node_modules/vba-extractor/scripts/Build-Xlsm.ps1 -SourceDir src/vba -OutputPath MyBook.xlsm

# .cls ファイル(と .bas の再同期)を、作成したワークブックに同期する
vba-extractor import MyBook.xlsm src/vba
```

> `-OutputPath` に相対パスを渡した場合、内部で `Workbook.SaveAs` を呼ぶ前にスクリプト実行時の
> カレントディレクトリ基準の絶対パスへ変換する（COM経由で起動したExcel自身の既定の作業フォルダ
> 基準で解決されてしまうのを防ぐため）。

詳細・注意点はスクリプト自身の `Get-Help -Full`(またはヘッダーコメント)を参照してください。

## 出力形式

| モジュール種別 | 拡張子 |
|---|---|
| 標準モジュール | `.bas` |
| クラスモジュール / フォーム / シートオブジェクト | `.cls` |

## 関連パッケージ

- [vba-runner](https://www.npmjs.com/package/vba-runner) — Excel 不要で VBA を実行・テストできるエンジン
