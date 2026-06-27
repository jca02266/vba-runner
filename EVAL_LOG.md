# vba-runner 評価ログ

サブエージェントによる「新規ユーザー視点での使い勝手評価」の累積記録。
**今後の評価を担当するサブエージェントはこのファイルを先に読み、過去に実施済みの
テストケースと発見済みの問題を把握した上で、新しい観点から評価を行うこと。**

---

## 評価済みドメイン・機能

| # | ドメイン | 主にテストした機能 | 日付 |
|---|---|---|---|
| 1 | 在庫管理システム | `.cls` クラス / `ReDim Preserve` / `On Error GoTo`・`Resume Next` / `Err.Raise` / 動的配列 | 2026-06-26 |
| 2 | ローマ数字コンバーター | `.cls` / `Property Get` / `ByRef` writeback / Boolean 変換 / JS 配列→VBA 配列 | 2026-06-27 |
| 3 | テキスト統計アナライザー | `Function As Double` 精度 / ディレクトリ読み込み / `eval()` 括弧あり・なし呼び出し / `Err.Raise` | 2026-06-27 |
| 4 | 図書館蔵書管理システム | `Scripting.Dictionary`（Add/Item/Exists/Count/Keys/Items/For Each/ネスト） / `VBA Collection`（Add/Item(1-based)/Item(key)/Count/Remove/For Each） / Dictionary+Collection 組み合わせ / クラス (`Book.cls`) / `On Error GoTo` | 2026-06-27 |

---

## 発見した問題と対応状況

### 修正済みバグ

| 問題 | 最小再現コード | 修正コミット |
|---|---|---|
| `eval()` で組み込み関数戻り値への `+`/`-` 演算が Error 424 | `r.eval('UBound(arr) + 1')` → Error 424（括弧ワークアラウンド: `(UBound(arr)) + 1`）| `ec63519` |
| `run()` ログで JS 配列引数が `[Object]` と表示される | `r.run('Proc', [[1,2,3]])` → ログが `Proc([Object])` | `ec63519` |
| `Dictionary.Item("nonexistent")` がキーを自動生成しない | 実 VBA では存在しないキーへの `.Item` 読み取りで Empty のエントリを自動生成する（Count+1, Exists→True）。修正後は VBA 互換動作＋コンソール警告を出力 | `ca409b7` |

### 仕様準拠の動作（バグではない）

| 現象 | 説明 |
|---|---|
| `eval('m + 1')`（m は Long 変数）→ Error 424 | 実 VBA でも変数名 + 算術は文として `Call m(+1)` と解釈される。ワークアラウンド: `(m) + 1` |
| `eval('Dim x : x = 42 : x')` → `undefined` | マルチステートメント末尾の裸の識別子は Call 文扱い。値を得るには別の `eval('x')` で |

### 未対応の機能制限（改善候補）

| 制限 | 詳細 |
|---|---|
| ~~`VBARunner` が複数ファイルの配列渡しに非対応~~ | **対応済み**（`new VBARunner(['/a/M1.bas', '/b/C1.cls'])` が動作するよう修正） |
| `eval()` の行番号が常に "line 1" | マルチライン eval でエラーが出ても行情報が `(line 1)` のみ |
| README に `eval()` の「式 vs 文」の注意書きがない | `eval('m + 1')` がエラーになる理由・`(m) + 1` の回避策が未記載 |
| `Dictionary.Add` へ Object をキーとして渡してもエラーなし | 実 VBA では非文字列キーの挙動は Object の hash になるが、引数順序ミス（Collection をキーに渡す）を検出できない。エラーなく格納されるが文字列で取り出せないため診断が困難 |

---

## 未テスト・未探索の領域（今後の評価で優先すること）

評価済みドメインでカバーしたものを除いた、まだ十分に試されていない機能。

### ファイル入出力

- `Open ... For Output/Input/Append`, `Print #`, `Line Input #`, `Close`
- Sandbox パス変換（`C:\` → サブディレクトリ変換）の動作確認
- 存在しないファイルの `Open For Input` → Error 53 (File not found)

### Scripting.Dictionary

- ~~`CreateObject("Scripting.Dictionary")` / `.Add` / `.Item` / `.Exists` / `.Keys` / `.Items`~~ **評価済み（評価#4）**
- ~~ネストした Dictionary~~ **評価済み（評価#4）**
- ~~`For Each` による Keys 列挙~~ **評価済み（評価#4）**

### 文字列操作の深掘り

- `Format()` 関数（数値フォーマット `"#,##0.00"` / 日付フォーマット）
- `InStr` / `InStrRev` / `Split` / `Join` の境界ケース
- 全角・マルチバイト文字の `Len` / `Mid` / `Left` / `Right`

### 数値型の境界

- `Integer`（-32768〜32767）のオーバーフロー → Error 6
- `Long` / `Double` / `Currency` の型強制と精度
- `CDbl` / `CLng` / `CInt` などの変換関数

### コレクション

- ~~VBA `Collection` オブジェクト（`Add` / `Item` / `Count` / `Remove` / `For Each`）~~ **評価済み（評価#4）**
- ~~ゼロ基底 vs 1基底のインデックスの挙動~~ **評価済み（評価#4）: 1基底が正しく動作する**

### 複数クラスの連携

- クラス間の相互参照（A が B のインスタンスを持つ）
- `Set` 代入 / `Is Nothing` 判定
- `Class_Terminate` の呼ばれるタイミング

### モック機能

- `__mocks__` ディレクトリによる VBA/JS/TS モック注入
- `setBuiltinOverride` で組み込み関数を上書き

### 条件付きコンパイル

- `#If` / `#Const` / `#Else`
- `config.compilerConstants` による定数上書き

### エラーハンドリングの深掘り

- `Resume` / `Resume Next` / `Resume Label`
- ネストした `On Error` （Sub から呼んだ Sub でエラーが起きた場合）
- `Err.Clear` / `Err.Number` の伝搬

### `VBARunner.run()` の高度な使い方

- `type: 'get'` / `'let'` / `'set'` オプション（Property アクセス）
- `ByRef` パラメーター経由での値の書き戻し確認
- `config.env` / `config.sandboxRoot` オプション

---

## 評価時の注意事項（過去の評価で学んだこと）

1. **配列でファイルを渡せる**（修正済み）: `new VBARunner(['/a/M1.bas', '/b/C1.cls'])` が動作する。なお配列渡しは mock スキャンが行われない（ディレクトリ渡しと異なる点）
2. **`eval()` で変数 + 算術は括弧が必要**: `eval('x + 1')` はエラー。`eval('(x) + 1')` または `eval('x')` で値を取り出してから JS 側で計算する
3. **組み込み関数の戻り値への算術は括弧不要（修正済み）**: `eval('UBound(arr) + 1')` は `ec63519` 以降正常に動作する
4. **ディレクトリ読み込みは `.bas`/`.cls` のみ**: `.ts` / `.js` ファイルが混在していても無視される
5. **`eval()` 末尾の裸の識別子は値を返さない**: マルチステートメント中の最後の `x` は `undefined`。値読み出しは独立した `eval('x')` で行う
6. **`Dictionary.Add(key, item)` vs `Collection.Add(item [, key])`**: 引数順序が逆。Dictionary はキーが先、Collection はアイテムが先。間違えると Error 91 が出るが「何を間違えたか」のヒントがない
7. **`Dictionary.Item("nonexistent")` はキーを自動生成する（VBA 互換）**: 読み取り時に存在しないキーがあると Empty でエントリを生成しコンソール警告を出力する
