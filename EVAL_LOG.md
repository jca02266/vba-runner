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
| 5 | CSV ログ書き込み・読み込みシステム | `Open For Output/Append/Input` / `Print #` / `Write #` / `Line Input #` / `Input #`（CSV）/ `Close` / `EOF()` / `FreeFile()` / `LOF()` / `LOC()` / Windows パス→VFS マッピング（C:\, D:\ ドライブ）/ 相対パス / `sandboxRoot` オプション / Error 53 / `Scripting.FileSystemObject`（CreateTextFile, OpenTextFile, FileExists, TextStream.ReadLine/ReadAll/WriteLine/Close）/ `Tab()` / VFS 事前配置 | 2026-06-27 |
| 6 | 家計簿・収支管理システム | `Integer`/`Long` オーバーフロー（Error 6）/ `Currency` 型の精度（浮動小数点のまま・要注意）/ `CInt`/`CLng`/`CCur` 変換関数・バンカーズ丸め / `Format()` `"#,##0.00"`・`"0.00%"` 正常・**`"000"` 零埋めバグ** / `InStr`/`InStrRev`（境界・開始位置・大文字小文字）/ `Split`/`Join`（空文字列・デリミタ）/ 全角文字の `Len`/`Mid`/`Left`/`Right`（文字数カウント正常）/ `On Error GoTo` / `Collection` + `Property Get/Let` クラス / `Debug.Print` 出力 | 2026-06-27 |
| 7 | 診療予約管理システム | 複数クラス連携（`Patient.cls` + `Appointment.cls`）/ `Set` 代入 / `Is Nothing` / `Class_Terminate` タイミング / 日付リテラル `#yyyy/mm/dd#` / `Format()` 日付パターン全般 / `DateSerial` / `DateAdd` / `DateDiff` / `Year`/`Month`/`Day` / `Now()` / `Date()` / `CDate` / `DateValue` / `IsDate` / `Weekday` / 日付+時刻リテラル `#yyyy/mm/dd HH:MM:SS#` | 2026-06-28 |
| 8 | バリデーション付き設定ファイルローダー | `Resume` / `Resume Next` / `Resume Label`（ラベルジャンプ）/ 複数スタックフレームを超えたエラー伝搬 / `Err.Clear` / `Err.Number` / `run()` type:'get','let','set' / JS モックオブジェクト Property Set 注入 / `config.env` + `Environ()` 注入 / `config.sandboxRoot` カスタム VFS ルート / `ByRef` 複数パラメーター writeback / `#If`/`#Const`/`#Else`/`#End If` 条件付きコンパイル / `config.compilerConstants` 外部定数注入 | 2026-07-03 |
| 9 | VS Code 拡張 LSP 機能（直接インポートによる評価） | メンバー補完（Dictionary/Worksheet/ユーザー定義クラス）/ チェーンアクセス引数なし `ws.Cells.` → Range / CreateObject 型推論 / VBA016 診断 / ホバー情報 `getMemberHoverInfo` | 2026-07-04 |
| 9 | VS Code 拡張 LSP 機能（初回評価） | `CompletionProvider.getCompletions` / `detectMemberAccess` / `resolveExprType` / `getMemberHoverInfo` / `checkUnknownTypes` / `collectUserDefinedTypeNames` / 単純メンバー補完（`dict.`/`ws.`）/ チェーン補完（`ws.Cells.` → Range）/ 引数付きチェーン `ws.Cells(1,1).`（バグ）/ ユーザー定義クラス補完 / `createObject` ProgID 型推論 / VBA016 未知型診断（column ずれバグ）/ mid-word ホバー（正常） | 2026-07-04 |

---

## 発見した問題と対応状況

### 修正済みバグ

| 問題 | 最小再現コード | 修正コミット |
|---|---|---|
| `eval()` で組み込み関数戻り値への `+`/`-` 演算が Error 424 | `r.eval('UBound(arr) + 1')` → Error 424（括弧ワークアラウンド: `(UBound(arr)) + 1`）| `ec63519` |
| `run()` ログで JS 配列引数が `[Object]` と表示される | `r.run('Proc', [[1,2,3]])` → ログが `Proc([Object])` | `ec63519` |
| `Dictionary.Item("nonexistent")` がキーを自動生成しない | 実 VBA では存在しないキーへの `.Item` 読み取りで Empty のエントリを自動生成する（Count+1, Exists→True）。修正後は VBA 互換動作＋コンソール警告を出力 | `ca409b7` |
| `Write #` で Boolean が `#TRUE#`/`#FALSE#` でなく `True`/`False` になる | `evaluateWriteStatement` に `VbaBoolean` 分岐を追加 | `9e25adc` |
| FSO `TextStream.ReadAll()` が `ReadLine()` 後も全体を返す | `readall` が `pos` を参照するよう修正 | `9e25adc` |
| `eval("Exit Sub")` が JS 例外を漏らしてクラッシュ | `executeStatements` を try/catch でラップして Exit シグナルを飲み込む | `0ca97d8` |
| 同一ファイルへの二重 `Open` が Error 55 を出さない | `fileHandles` を走査して同一パスの重複チェックを追加 | `0ca97d8` |
| FSO `TextStream.AtEndOfStream` 未実装（Error 438） | `pos >= content.length` を返す getter を実装 | `0ca97d8` |

### 仕様準拠の動作（バグではない）

| 現象 | 説明 |
|---|---|
| `eval('m + 1')`（m は Long 変数）→ Error 424 | 実 VBA でも変数名 + 算術は文として `Call m(+1)` と解釈される。ワークアラウンド: `(m) + 1` |
| `eval('Dim x : x = 42 : x')` → `undefined` | マルチステートメント末尾の裸の識別子は Call 文扱い。値を得るには別の `eval('x')` で |
| `eval()` で他モジュールの `Dim`/`Private` 変数が読めない | `eval()` は独立したトップレベルモジュールとして評価されるため、他モジュールの `Private` 変数にはアクセスできない（VBA のクロスモジュール非公開変数と同じ意味論）。`Public` 変数はグローバル env 経由でアクセス可能。意図通りの設計。 |

### 拡張機能 LSP のバグ（評価 #9 で発見・未修正）

| 問題 | 再現コード | 根本原因 |
|---|---|---|
| 引数付きチェーン補完が効かない | `ws.Cells(1, 1).` で Range でなくグローバル関数 29 件が返る | `detectMemberAccess` の正規表現が `)` 終端の式にマッチしない（`completion-provider.ts:454`） |
| VBA016 波下線が変数名を指す | `Dim x As UnknownType` で `x` に波下線（`UnknownType` であるべき） | `unknown-type-checker.ts:79` で `d.name.loc`（変数名位置）を渡している。パーサーの `VariableDeclarator` に型名の loc フィールドなし |

### 未対応の機能制限（改善候補）

| 制限 | 詳細 |
|---|---|
| ~~`VBARunner` が複数ファイルの配列渡しに非対応~~ | **対応済み**（`new VBARunner(['/a/M1.bas', '/b/C1.cls'])` が動作するよう修正） |
| `eval()` の行番号が常に "line 1" | マルチライン eval でエラーが出ても行情報が `(line 1)` のみ |
| README に `eval()` の「式 vs 文」の注意書きがない | `eval('m + 1')` がエラーになる理由・`(m) + 1` の回避策が未記載 |
| `Dictionary.Add` へ Object をキーとして渡してもエラーなし | 実 VBA では非文字列キーの挙動は Object の hash になるが、引数順序ミス（Collection をキーに渡す）を検出できない。エラーなく格納されるが文字列で取り出せないため診断が困難 |
| ~~`Exit Sub` を `eval()` トップレベルで使うと JS 例外が漏れる~~ | **修正済み** (`0ca97d8`): `executeStatements` を try/catch でラップして Exit シグナルを飲み込む |
| ~~`Write #` で Boolean が `#TRUE#`/`#FALSE#` でなく `True`/`False` になる~~ | **修正済み** (`9e25adc`): `evaluateWriteStatement` に `VbaBoolean` 分岐を追加 |
| ~~`FSO TextStream.AtEndOfStream` 未実装~~ | **修正済み** (`0ca97d8`): `pos >= content.length` を返す getter を実装 |
| ~~`FSO TextStream` 位置追跡バグ~~ | **修正済み** (`9e25adc`): `readall` が `pos` を参照するよう修正 |
| ~~同一ファイルへの二重 `Open` が Error 55 を出さない~~ | **修正済み** (`0ca97d8`): `fileHandles` を走査して同一パスの重複チェックを追加 |
| ~~`Format()` の零埋めが動作しない~~ | **修正済み**: `intPart.padStart(minIntegers, '0')` を追加。`Format(42, "000")` → `"042"` が正常動作 |
| `Currency` 型が固定小数点演算でない | 実 VBA では `Currency` は 4 桁固定小数点（0.1+0.2 = 0.3 厳密）。vba-runner では `Double` と同じ浮動小数点演算になる（0.1+0.2 = 0.30000000000000004）。 |
| LSP: 引数付きチェーン補完 `obj.Method(args).` が効かない | `detectMemberAccess` の正規表現 `/([a-zA-Z_][a-zA-Z0-9_]*)\.([a-zA-Z_][a-zA-Z0-9_]*)?$/` は末尾が識別子の場合のみマッチし、`ws.Cells(1,1).` のように `)` で終わる式はスキップする。その結果チェーン解決コードが実行されず、標準関数一覧が返る。`ws.Cells.`（引数なし）は正常動作（48件 Range メンバー）。`src/lsp/completion-provider.ts` line 454 の正規表現を `[a-zA-Z_][a-zA-Z0-9_]*|\([^)]*\)` のような形式に拡張が必要。 |
| LSP: VBA016 診断の range が型名でなく変数名を指す | `unknown-type-checker.ts` line 79: `warn(d.objectType, d.name.loc ?? stmt.loc)` で `d.name.loc` は変数名（`x`）の位置を渡している。型名（`UnknownType`）の位置ではないため、エディター上の波下線が変数名に表示される（col 4、期待は col 9）。パーサー側に `objectTypeLoc` フィールドを追加するか、warn 内でオフセット補正が必要。最小再現: `Dim x As UnknownType` で診断 range が col=4 (変数名 `x`) を指す。 |
| `Dim empty As String` がパースエラー | `empty` は VBA 仕様上の予約語のため変数名に使えない。VBA 仕様準拠の正しい動作。 |
| ~~時刻のみの日付リテラル未対応~~ | **修正済み** (`b4d00c3`): `#12:30:45#` / `#8:30:00 AM#` が Error 13 になっていた。`parseDateLiteral` が時刻のみの場合に基準日（1899/12/30）を返すよう修正。 |
| ~~`Class_Terminate` が参照カウントなしで早期発動~~ | **修正済み**: `Set p1 = Nothing` を呼んでも別変数 `p2` が同じオブジェクトを保持していれば `Class_Terminate` を呼ばないよう、参照カウント（`__refCount__`）を実装。`Set` 代入で addRef、`Set = Nothing` で releaseRef、スコープ脱出でもカウントを減算。`circular-reference-terminate.test.ts` 全 16 テスト通過。 |

---

## 未テスト・未探索の領域（今後の評価で優先すること）

評価済みドメインでカバーしたものを除いた、まだ十分に試されていない機能。

### ファイル入出力

- ~~`Open ... For Output/Input/Append`, `Print #`, `Line Input #`, `Close`~~ **評価済み（評価#5）**
- ~~Sandbox パス変換（`C:\` → サブディレクトリ変換）の動作確認~~ **評価済み（評価#5）: 正常動作**
- ~~存在しないファイルの `Open For Input` → Error 53 (File not found)~~ **評価済み（評価#5）: 正常動作**

### Scripting.Dictionary

- ~~`CreateObject("Scripting.Dictionary")` / `.Add` / `.Item` / `.Exists` / `.Keys` / `.Items`~~ **評価済み（評価#4）**
- ~~ネストした Dictionary~~ **評価済み（評価#4）**
- ~~`For Each` による Keys 列挙~~ **評価済み（評価#4）**

### 文字列操作の深掘り

- ~~`Format()` 関数（数値フォーマット `"#,##0.00"` / `"0.00%"`）~~ **評価済み（評価#6）: `"#,##0.00"` / `"0.00%"` は正常。`"000"` 零埋めはバグあり**
- ~~`InStr` / `InStrRev` / `Split` / `Join` の境界ケース~~ **評価済み（評価#6）: 正常動作**
- ~~全角・マルチバイト文字の `Len` / `Mid` / `Left` / `Right`~~ **評価済み（評価#6）: 文字数カウントで正常動作**
- ~~`Format()` 日付フォーマット（`"yyyy/mm/dd"` 等）— 日付リテラル `#2024/01/15#` の扱い未確認~~ **評価済み（評価#7）: 日付リテラル / Format 日付パターン / DateAdd / DateDiff / Now / DateSerial 等すべて正常動作。時刻のみリテラル `#HH:MM:SS#` は未対応（バグ）**

### 数値型の境界

- ~~`Integer`（-32768〜32767）のオーバーフロー → Error 6~~ **評価済み（評価#6）: 正常動作**
- ~~`Long`（-2147483648〜2147483647）のオーバーフロー → Error 6~~ **評価済み（評価#6）: 正常動作**
- ~~`CInt` / `CLng` / `CCur` などの変換関数~~ **評価済み（評価#6）: バンカーズ丸めも正常**
- ~~`Currency` 型の精度~~ **評価済み（評価#6）: 浮動小数点扱いのため 0.1+0.2≠0.3（バグ）**
- `CDbl` / `CSng` の精度と `Single` 型の動作

### コレクション

- ~~VBA `Collection` オブジェクト（`Add` / `Item` / `Count` / `Remove` / `For Each`）~~ **評価済み（評価#4）**
- ~~ゼロ基底 vs 1基底のインデックスの挙動~~ **評価済み（評価#4）: 1基底が正しく動作する**

### 複数クラスの連携

- ~~クラス間の相互参照（A が B のインスタンスを持つ）~~ **評価済み（評価#7）: 正常動作**
- ~~`Set` 代入 / `Is Nothing` 判定~~ **評価済み（評価#7）: 正常動作**
- ~~`Class_Terminate` の呼ばれるタイミング~~ **評価済み（評価#7）＋修正済み: 参照カウント実装により最後の参照が解放されたときに発動するよう修正**

### 拡張機能 LSP（未テスト）

- `__mocks__` ディレクトリによる VBA/JS/TS モック注入
- `setBuiltinOverride` で組み込み関数を上書き
- vba-types.json の自動リロード（FileSystemWatcher）
- With ブロック内のチェーン補完（`.Cells.` など）
- クロスモジュール補完（複数ファイル展開時）
- VBA016 Quick Fix の動作（initTypeStubs / addToTypeStubs コマンド）

### ~~条件付きコンパイル~~ **評価済み（評価#8）**

- ~~`#If` / `#Const` / `#Else`~~ **評価済み（評価#8）: 正常動作**
- ~~`config.compilerConstants` による定数上書き~~ **評価済み（評価#8）: ただしファイル内 `#Const` が `compilerConstants` より優先（VBA 仕様準拠）。VBA7 デフォルト=0 に注意**

### ~~エラーハンドリングの深掘り~~ **評価済み（評価#8）**

- ~~`Resume` / `Resume Next` / `Resume Label`~~ **評価済み（評価#8）: 全3形式とも正常動作。`Resume`（ラベルなし）は同一行リトライ確認**
- ~~ネストした `On Error` （Sub から呼んだ Sub でエラーが起きた場合）~~ **評価済み（評価#8）: 複数スタックフレームを超えた伝搬が正常動作（MidLevel 無ハンドラー → TopLevel でキャッチ確認）**
- ~~`Err.Clear` / `Err.Number` の伝搬~~ **評価済み（評価#8）: `Err.Clear` で `Err.Number` が 0 にリセットされることを確認**

### ~~`VBARunner.run()` の高度な使い方~~ **評価済み（評価#8）**

- ~~`type: 'get'` / `'let'` / `'set'` オプション（Property アクセス）~~ **評価済み（評価#8）: `type:'set'` で JS モックオブジェクトを VBA Property Set へ注入可能（依存性注入パターンとして有用）**
- ~~`ByRef` パラメーター経由での値の書き戻し確認~~ **評価済み（評価#8）: 複数 ByRef パラメーターすべて JS 配列への writeback を確認**
- ~~`config.env` / `config.sandboxRoot` オプション~~ **評価済み（評価#8）: 両オプションとも正常動作**

---

## 評価時の注意事項（過去の評価で学んだこと）

1. **配列でファイルを渡せる**（修正済み）: `new VBARunner(['/a/M1.bas', '/b/C1.cls'])` が動作する。なお配列渡しは mock スキャンが行われない（ディレクトリ渡しと異なる点）
2. **`eval()` で変数 + 算術は括弧が必要**: `eval('x + 1')` はエラー。`eval('(x) + 1')` または `eval('x')` で値を取り出してから JS 側で計算する
3. **組み込み関数の戻り値への算術は括弧不要（修正済み）**: `eval('UBound(arr) + 1')` は `ec63519` 以降正常に動作する
4. **ディレクトリ読み込みは `.bas`/`.cls` のみ**: `.ts` / `.js` ファイルが混在していても無視される
5. **`eval()` 末尾の裸の識別子は値を返さない**: マルチステートメント中の最後の `x` は `undefined`。値読み出しは独立した `eval('x')` で行う
6. **`Dictionary.Add(key, item)` vs `Collection.Add(item [, key])`**: 引数順序が逆。Dictionary はキーが先、Collection はアイテムが先。間違えると Error 91 が出るが「何を間違えたか」のヒントがない
7. **`Dictionary.Item("nonexistent")` はキーを自動生成する（VBA 互換）**: 読み取り時に存在しないキーがあると Empty でエントリを生成しコンソール警告を出力する
8. **VFS パスは絶対パス `/sandbox/c/...` を使う**: `sandboxRoot` はコンストラクターオプション名。JS から VFS へアクセスする際は先頭スラッシュが必須（`fs.readFileSync('/sandbox/c/test.txt', 'utf-8')`）
9. **`Print #` は CRLF（`\r\n`）を書く**: JS 側で `readFileSync` したときは `split('\r\n')` でパースすること。`split('\n')` では行末に `\r` が残る
10. **~~`Exit Sub` を `eval()` トップレベルで呼ばない~~（修正済み `0ca97d8`）**: 現在は `executeStatements` が Exit シグナルを飲み込むため正常動作する
11. **~~FSO `AtEndOfStream` は未実装（Error 438）~~（修正済み `0ca97d8`）**: 現在は `pos >= content.length` を返す getter が実装済み。`Do While Not ts.AtEndOfStream` パターンも使用可
12. **`Format()` の零埋め（`"000"` 等）は正常動作**（修正済み）: `Format(42, "000")` → `"042"`。`"#,##0.00"` / `"0.00%"` も正常。
13. **`Currency` は固定小数点ではなく浮動小数点**: `Dim c As Currency: c = 0.1 + 0.2` は `Double` と同じ誤差が出る。精度が必要な場合は JS 側で計算するか `Round()` を使う。
14. **`empty` / `Empty` は変数名に使えない**: `Dim empty() As String` は「Expected variable name (Found empty)」でパースエラー。`emptyArr` 等の代替名を使う。
15. **`run()` の第2引数は省略可能**（修正済み）: 引数なし Sub は `r.run('Sub名')` と省略できる。`args` のデフォルト値を `[]` に修正済み。
16. **時刻のみリテラル `#HH:MM:SS#` は使えない**: `#12:30:45#` や `#8:30:00 AM#` は Error 13 Type mismatch。日付+時刻 `#2024/01/15 12:30:45#` はOK。
17. **`Class_Terminate` は参照カウントなしで早期発動する**: `Set p1 = Nothing` で他に参照があっても `Class_Terminate` が呼ばれる。ただしオブジェクト自体は破棄されず残存する。VBA の COM 参照カウント完全再現ではない（evaluator.ts に既知制限として明記）。
18. **`Format()` 日付パターンは豊富に動作する**: `yyyy/mm/dd` / `yyyy年mm月dd日` / `d-mmm-yyyy` / `dddd` / `ddd` / `Long Date` / `Short Date` / `yy/mm/dd` / `hh:mm:ss` / `h:mm AM/PM` すべて正常。
19. **`.cls` ファイルのクラス名はファイル名（拡張子なし）で決まる**: `Attribute VB_Name = "MyObj"` の値ではなく、`MyObj.cls` のようにファイル名がクラス名になる。ファイルを `TerminateTest.cls` と名付けると VBA 側で `New TerminateTest` と書かないと Error 429 になる。実際の VBA エクスポートではファイル名と `VB_Name` は通常一致しているが、ファイルをリネームした場合に落とし穴になる。
20. **コンパイラ定数のデフォルトは現代的な 64bit Windows 環境**: `DEFAULT_COMPILER_CONSTANTS: { VBA7: -1, Win64: -1, Win32: -1, Mac: 0 }`（Office 2010+ / 64bit）。32bit 環境をシミュレートするには `compilerConstants: { VBA7: 0, Win64: 0 }` を渡す。README のセクション10に記載。
21. **ファイル内 `#Const` は外部 `compilerConstants` より優先**: `preprocess()` は `localConsts`（`#Const` で定義）を `merged`（`compilerConstants`）より先に参照する。外部から定数を注入しても、ファイル内で `#Const FOO = X` が定義されていれば上書きされる。VBA 仕様準拠。README のセクション10に記載。
22. **`Dim entry As New Type` をループ内で使うと同一オブジェクトを共有する**:
 VBA の `Dim` はプロシージャスコープのため、ループ内に書いてもループ変数はイテレーション間で共有される。新しいオブジェクトが必要なら `Set entry = New Type` をループ内に書くこと（`Dim` はプロシージャ先頭で1回だけ宣言する）。
23. **`run()` の `type:'set'` でJS モックオブジェクトを Property Set へ注入できる**: `r.run('PropName', [mockObj], 'set')` でフラットな JS オブジェクトを VBA の Property Set 経由でモジュール変数に代入できる。その後 VBA 側のコードで `Is Nothing` 判定や、プロパティアクセスが可能。依存性注入パターンとして有用。
24. **`eval()` で読めるのは `Public` モジュール変数のみ（`Dim`/`Private` は不可）**: `eval()` は独立したトップレベルモジュールとして評価されるため、他モジュールの `Private`/`Dim` 変数は見えない。`Public` 変数はグローバル env 経由でアクセス可能。意図通りの設計（モジュールコンテキスト内でのデバッグ評価とは別概念）。
25. **`CompletionProvider` を LSP 外部から使うには `Parser` を `errorRecovery: true` で呼ぶこと**: `new Parser(tokens).parse()` は不完全な VBA（補完トリガー時の途中入力）でスローする。LSP サーバーが内部で使っているように `new Parser(tokens, { errorRecovery: true }).parse()` とする必要がある。公開ドキュメントに記載なし。
26. **`CompletionProvider.getCompletions` の第1引数は `Program.body`（`Statement[]`）**: `parser.parse()` が返す `Program` オブジェクトをそのまま渡すと TypeError。`.body` を取り出して渡すこと。
27. **`checkUnknownTypes` の第2引数は `Set<string>`（オブジェクトリテラル `{}` は不可）**: 誤って `{}` を渡すと `knownTypeNames.has is not a function` で実行時エラー。`new Set<string>()` か `collectUserDefinedTypeNames(stmts)` の戻り値を渡すこと。
