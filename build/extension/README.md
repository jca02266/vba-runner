# VBA Runner — VS Code 拡張機能

VBA Runner の VS Code 拡張機能です。`.bas` / `.cls` / `.frm` ファイルを VS Code で開くと、LSP（言語サーバープロトコル）による豊富な編集支援が有効になります。

## インストール

`.vsix` ファイルをダウンロードして以下のいずれかの方法でインストールします。

```bash
# コマンドラインからインストール
code --install-extension vba-runner-0.1.0.vsix
```

または VS Code の「拡張機能」ビューで「...」メニュー → 「VSIXからインストール」を選択します。

**開発用（リポジトリを clone した場合）**:

1. このリポジトリを VS Code で開く
2. `F5` キーで「Extension Development Host」を起動する
3. 新しいウィンドウで `.bas` / `.cls` / `.frm` ファイルを開く

## 対応ファイル

| 拡張子 | 種別 |
|---|---|
| `.bas` | 標準モジュール |
| `.cls` | クラスモジュール |
| `.frm` | フォームモジュール |

## 実装済みLSP機能

### ホバー (Hover)

シンボル（Sub/Function/変数/定数/クラス/イベント）にマウスカーソルを合わせると、シグネチャをポップアップ表示します。

```vb
Sub CalcSum(a As Integer, b As Integer)
'         ↑ ここにホバー → "Sub CalcSum(a As Integer, b As Integer)" が表示される
```

### 定義へ移動 (Go to Definition)

シンボルにカーソルを置いて `F12` を押すと、そのシンボルの宣言位置にジャンプします。

| 対応シンボル | 例 |
|---|---|
| Sub / Function / Property | `Sub Foo()`, `Function Bar() As Long` |
| 変数宣言 | `Dim x As Integer`, `Public count As Long` |
| 定数 | `Const MAX As Long = 100` |
| クラス | `Class MyClass` |
| イベント | `Event DataChanged(newVal As Variant)` |

### コード補完 (Completion)

入力中に VBA キーワード・組み込み関数・ソース内で定義されたプロシージャの候補を表示します。`.` を入力したときにもトリガーされます。

### コール階層（Call Graph）

コマンドパレット（`Ctrl+Shift+P`）から以下のコマンドを実行します。

- **VBA: Show Call Graph** — カーソル位置のプロシージャを起点にした呼び出しグラフを表示
- **VBA: Show in Call Graph** — カーソル位置のプロシージャを全体のグラフ中でハイライト

### 変数の導入リファクタリング (Introduce Variable)

コマンドパレットから **Refactor: Introduce Variable** を実行すると、選択した式を変数に抽出するリファクタリングを実行します。

### VBA デバッガー統合

`.bas` ファイルを開いた状態で `F5` を押すと、VBA Runner のデバッガーでファイルを起動できます（`launch.json` の設定不要）。

## 設定オプション

| 設定 | デフォルト | 説明 |
|---|---|---|
| `vba-runner.lint.enabled` | `false` | lint 警告（VBA001〜 等のコード付き診断）をすべて有効にする |
| `vba-runner.lint.enabledCodes` | `[]` | 有効にする lint コードを個別指定する（例: `["VBA009"]`） |
| `vba-runner.editor.autoLineContinuation` | `true` | 式の途中で Enter を押したとき行継続文字 `_` を自動挿入する |

## ビルド

```bash
npm run build:extension   # dist/extension.cjs を生成
npm run package:extension # .vsix パッケージを生成
```

## 詳細ドキュメント

- [LSP.md](../../LSP.md) — LSP の設計・実装方針
- [REFERENCE.md](../../REFERENCE.md) — 詳細仕様
- [README.md](../../README.md) — プロジェクト概要
