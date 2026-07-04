# VBA Runner — VS Code 拡張機能

VBA Runner の VS Code 拡張機能です。`.bas` / `.cls` / `.frm` ファイルを VS Code で開くと、LSP（言語サーバープロトコル）による豊富な編集支援が有効になります。

## インストール

VS Code マーケットプレイスで **VBA Runner** を検索するか、以下のコマンドを実行します:

```bash
code --install-extension jca02266.vba-runner
```

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

## 編集支援（LSP）

### ホバー (Hover)

シンボル（Sub / Function / 変数 / 定数 / クラス / イベント）にマウスカーソルを合わせると、シグネチャをポップアップ表示します。

```vb
Sub CalcSum(a As Integer, b As Integer)
'         ↑ ここにホバー → "Sub CalcSum(a As Integer, b As Integer)" が表示される
```

### シグネチャヘルプ (Signature Help)

関数呼び出しの `(` または `,` を入力すると、パラメーター名と現在の引数位置をポップアップ表示します。VBA 組み込み関数（`MsgBox`・`Format` 等）と自分で定義した Sub / Function の両方に対応しています。

### 定義へ移動 (Go to Definition)

シンボルにカーソルを置いて `F12` を押すと、そのシンボルの宣言位置にジャンプします。

| 対応シンボル | 例 |
|---|---|
| Sub / Function / Property | `Sub Foo()`, `Function Bar() As Long` |
| 変数宣言 | `Dim x As Integer`, `Public count As Long` |
| 定数 | `Const MAX As Long = 100` |
| クラス | `Class MyClass` |
| イベント | `Event DataChanged(newVal As Variant)` |

### 参照の検索 (Find References)

シンボルにカーソルを置いて `Shift+F12` を押すと、そのシンボルが参照されているすべての箇所をリスト表示します。

### シンボルのリネーム (Rename Symbol)

シンボルにカーソルを置いて `F2` を押すと、ファイル内のすべての参照箇所をまとめてリネームします。

### コード補完 (Completion)

入力中に VBA キーワード・組み込み関数・ソース内で定義されたプロシージャの候補を表示します。

#### メンバー補完 (`.`)

変数に `.` を入力すると、宣言された型に固有のメンバーを表示します。対応型は以下の通りです:

| 型 | トリガー例 |
|---|---|
| `Scripting.Dictionary` | `Dim d As Scripting.Dictionary` → `d.` |
| `Scripting.FileSystemObject` | `Dim fso As Scripting.FileSystemObject` → `fso.` |
| `ADODB.Recordset` / `ADODB.Connection` | `Dim rs As ADODB.Recordset` → `rs.` |
| `RegExp` / `VBScript.RegExp` | `Dim re As RegExp` → `re.` |
| `Collection` | `Dim col As Collection` → `col.` |
| `Range` / `Worksheet` / `Workbook` | `Dim ws As Worksheet` → `ws.` |
| `Sheets` / `Application` | `Dim app As Application` → `app.` |
| ユーザー定義クラス | `Dim obj As MyClass` → `obj.` |

クロスモジュール補完にも対応しており、ワークスペース内の他の `.bas` / `.cls` ファイルで宣言されたクラスも自動的に認識されます。

#### チェーンアクセス解決

メンバーの戻り値型を追跡するため、連続した `.` 補完も正しく解決されます:

```vb
Dim ws As Worksheet
ws.Cells.         ' → Range のメンバーを補完（Cells の戻り値は Range）
ws.Range("A1").Offset(1, 0).  ' → Range のメンバー（Offset の戻り値は Range）
ws.Parent.        ' → Workbook のメンバー（Parent の戻り値は Workbook）
```

#### With ブロック内の補完

`With` ブロック内で行頭に `.` を入力すると、`With` の対象オブジェクトのメンバーを補完します:

```vb
With ws
    .Cells.       ' → Range のメンバー
    .Name         ' → Worksheet の Name プロパティ
End With
```

#### スニペット

| プレフィックス | 展開内容 |
|---|---|
| `fe` | `For Each ... In ... Next` |
| `for` | `For ... To ... Next` |
| `sc` | `Select Case ... Case Else ... End Select` |
| `if` | `If ... Then ... Else ... End If` |
| `oeg` | `On Error GoTo ... ErrHandler` パターン |
| `wi` | `With ... End With` |
| `sub` | `Sub ... End Sub` |
| `fn` | `Function ... End Function` |
| `do` | `Do While ... Loop` |
| `dim` | `Dim` 変数宣言 |
| `dict` | `Dim … As Object` + `Set … = CreateObject("Scripting.Dictionary")` |
| `fso` | `Dim … As Object` + `Set … = CreateObject("Scripting.FileSystemObject")` |
| `regex` | `Dim … As Object` + `Set … = CreateObject("VBScript.RegExp")` + Pattern / Global |
| `adors` | `Dim … As Object` + `Set … = CreateObject("ADODB.Recordset")` |
| `adocn` | `Dim … As Object` + `Set … = CreateObject("ADODB.Connection")` + ConnectionString |
| `cobj` | `Dim … As Object` + `Set … = CreateObject("ProgID")` (汎用) |

### ドキュメントシンボル（アウトライン）

アウトラインパネル（`Ctrl+Shift+O`）とワークスペースシンボル検索（`Ctrl+T`）で、VBA ファイル内のすべての Sub / Function / Property / クラスメンバーを一覧表示します。

`' --- Name ---` や `' === Name ===` 形式のセクション区切りコメントも Namespace シンボルとしてアウトラインに表示されるため、大きなモジュールのナビゲーションに役立ちます。

```vb
' ─── 初期化処理 ─────────────────────────────────────
Public Sub Initialize()
    ...
End Sub

' === データ処理 ===
Public Function Process(data) As Long
    ...
End Function
```

## 診断（Diagnostics）

以下のルールがリアルタイムで報告されます:

| コード | 重要度 | ルール | 条件 |
|---|---|---|---|
| — | エラー | 構文エラー | パーサーが検出した構文エラー |
| VBA001 | 警告 | `ByVal`/`ByRef` 未指定 | パラメーターに明示的な引数渡し修飾子がない |
| VBA009 | 警告 | デッドストア | 変数が代入されているが読み取られていない |
| VBA011 | ヒント | Range アクセス | `Sheets("name")` を型付き変数で代替すべき |
| VBA013 | 警告 | `Option Explicit` 未記述 | ファイルに `Option Explicit` がない |
| VBA014 | 警告 | 未使用変数 | 宣言されているが一度も参照されていない変数 |
| VBA016 | 警告 | 未知の型 | `Dim x As UnknownType` — 型が認識されない |

**クイックフィックス**が提供されるルール:

- **VBA013** — *「Option Explicit」を追加*: ファイル先頭に `Option Explicit` を自動挿入
- **VBA016** — *「TypeName」を vba-types.json に追加*: 型のプレースホルダーエントリーを追記
- **VBA016** — *vba-types.json をすべての COM 型定義で初期化*: ファイルが存在しない場合のみ表示。組み込み COM 型をすべて含む初期ファイルを作成

lint ルール（VBA001・VBA009・VBA014 等）はデフォルトで無効です。設定で有効化します:

```json
{ "vba-runner.lint.enabled": true }
```

個別に有効化する場合:

```json
{ "vba-runner.lint.enabledCodes": ["VBA009", "VBA014"] }
```

## 外部型定義ファイル（`vba-types.json`）

拡張機能に組み込まれていない型（独自 COM オブジェクト・モッククラス・未対応の Excel 型など）のメンバー補完を追加するには、ワークスペースルートに `vba-types.json` を作成します:

```json
{
  "MyComObject": [
    { "label": "DoWork",  "kind": "Method",   "detail": "DoWork(arg As String) As Boolean" },
    { "label": "Status",  "kind": "Property", "detail": "Status As Long" }
  ],
  "MyHelper": [
    { "label": "Compute", "kind": "Function", "detail": "Compute(x As Long) As Double", "returnType": "myresult" }
  ]
}
```

**フィールド:**

| フィールド | 値 | 説明 |
|---|---|---|
| `label` | 文字列 | メンバー名 |
| `kind` | `"Method"` / `"Function"` / `"Property"` / `"Variable"` / `"Constant"` | アイコンとカテゴリー |
| `detail` | 文字列 | 補完ポップアップに表示するシグネチャ |
| `returnType` | 文字列（小文字の型名） | チェーンアクセス解決のための戻り値型 |

**クイックフィックスによる初期化:**

VBA016 診断が表示されたとき、クイックフィックスメニューから:

- *「TypeName」を vba-types.json に追加* — その型のプレースホルダーエントリーを追記
- *vba-types.json をすべての COM 型定義で初期化* — Scripting.Dictionary・Range・Worksheet 等を含む初期ファイルを作成（ファイルが存在しない場合のみ表示）

`vba-types.json` のエントリーは組み込み定義より**優先**されるため、組み込み型のメンバーリストを独自定義で上書きすることもできます。

ファイルの変更は自動的に反映されます。

## Code Lens

各プロシージャの宣言行にインラインボタンが表示されます。

| ボタン | 動作 |
|---|---|
| `▶ Run` | そのプロシージャを VBA Runner で実行 |
| `🐛 Debug` | デバッガーでステップ実行 |
| `N references` | 参照箇所を一覧表示 |
| `未テスト` / `✓ テスト済み` | テストスタブを生成 / テスト関数へジャンプ |
| `📊 Show in Call Graph` | コールグラフ上でハイライト |
| `✓ Nms` | テスト成功（例: `✓ 3ms`）— テスト実行後に表示 |
| `✗ メッセージ` | テスト失敗とエラーメッセージの先頭行 |

## フォーマット（整形）

`Shift+Alt+F`（または右クリック → **ドキュメントのフォーマット**）でファイルを自動整形します。`"editor.formatOnSave": true` と組み合わせることで保存時自動整形も可能です。

整形内容:
- `Sub` / `Function` / `If` / `For` / `With` / `Select Case` の一貫したインデント
- `Case` ラベルを `Select Case` と同じ列に揃える
- キーワードの大文字小文字を標準 VBA スタイルに統一

## リファクタリング

コマンドパレット（`Ctrl+Shift+P`）または Code Lens から実行できます。

| コマンド | 説明 |
|---|---|
| **Refactor: Introduce Variable** | 選択した式を変数に抽出 |
| **Refactor: Extract Function** | 選択した処理を新しいプロシージャに抽出 |
| **Refactor: Extract Constant** | 選択したリテラルを定数に抽出 |
| **Refactor: Inline Variable** | 変数をその使用箇所にインライン展開 |
| **Refactor: Introduce With** | オブジェクト参照を `With` ブロックにまとめる |
| **Refactor: Remove Unused Variables** | 未使用変数の宣言を削除 |
| **Refactor: Organize Declarations** | 変数宣言をプロシージャ先頭にまとめる |

## コール階層（Call Graph）

コマンドパレットから以下のコマンドを実行します。

- **VBA: Show Call Graph** — カーソル位置のプロシージャを起点にした呼び出しグラフを表示
- **VBA: Show in Call Graph** — カーソル位置のプロシージャを全体のグラフ中でハイライト

## テスト支援

### テストスタブ生成

Code Lens の「未テスト」ボタンを押すと、`Test_<プロシージャ名>` のスタブを生成します。初回実行時にテストの配置場所を選択します。

- **同じファイルに追加** — 同じ `.bas` ファイルの末尾に追記
- **別ファイルに追加** — `<ファイル名>Test.bas` を作成して追記

選択はワークスペース設定（`vba-runner.test.location`）に保存されます。

### テスト結果のインライン表示

テストプロシージャの **▶ Run** Code Lens でテストを実行すると、結果がインラインに表示されます:

- `✓ 3ms` — 3 ms でテスト成功
- `✗ Expected 1 but got 2` — アサーションメッセージの先頭行と共にテスト失敗

### モックひな形生成

コマンドパレットから **VBA: Generate Mocks** を実行すると、ソースファイルの Excel 依存オブジェクト（`Worksheet`・`Range` 等）を解析し、`__mocks__/ExcelObjects.bas` にモックのひな形を生成します。

## VBA デバッガー統合

`.bas` ファイルを開いた状態で `F5` を押すと、VBA Runner のデバッガーでファイルを起動できます（`launch.json` の設定不要）。

### `Debug.Print` の出力先

`Debug.Print` の出力は専用の **VBA Debug** Output Channel に表示されます（VBA Runner ログとは別）。出力が発生すると自動的にパネルが表示されます。

## 設定オプション

| 設定 | デフォルト | 説明 |
|---|---|---|
| `vba-runner.lint.enabled` | `false` | lint 警告（VBA001〜 等のコード付き診断）をすべて有効にする |
| `vba-runner.lint.enabledCodes` | `[]` | 有効にする lint コードを個別指定する（例: `["VBA009"]`） |
| `vba-runner.editor.autoLineContinuation` | `true` | 式の途中で Enter を押したとき行継続文字 `_` を自動挿入する |
| `vba-runner.editor.autoKeywordCasing` | `true` | キーワードの大文字小文字を確定時に自動補正する（VBE の動作に合わせる） |
| `vba-runner.test.location` | _(未設定)_ | テストスタブの配置先（`sameFile` / `separateFile`）。未設定時は初回実行時に選択 |

## ビルド

```bash
npm run build:extension   # dist/extension.cjs を生成
npm run package:extension # .vsix パッケージを生成
```

## 詳細ドキュメント

- [LSP.md](../../LSP.md) — LSP の設計・実装方針
- [REFERENCE.md](../../REFERENCE.md) — 詳細仕様
- [README.md](../../README.md) — プロジェクト概要
