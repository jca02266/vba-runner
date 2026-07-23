# vba-runner 評価ログ

サブエージェントによる「新規ユーザー視点での使い勝手評価」の累積記録。
**今後の評価を担当するサブエージェントはこのファイルを先に読み、過去に実施済みの
テストケースと発見済みの問題を把握した上で、新しい観点から評価を行うこと。**

---

## 評価済みドメイン・機能

| # | ドメイン | 主にテストした機能 | 日付 |
|---|---|---|---|
| 56 | 実行互換性 #56: サフィックス付き手続き名 | `Function Greeting$()` の戻り値代入と、サフィックスなしの `Greeting()` 呼び出しを確認。宣言名のサフィックスを戻り型へ反映して通常名に正規化する既存実装により、期待どおり文字列を返した。回帰テストを追加。 | 2026-07-23 |
| 55 | 実行互換性 #55: 修飾 `Implements` 名 | **Bug 55-A 修正済み**: `Implements ADODB.ICommand` が最初の識別子だけを読み、`.` で構文エラーになった。ドット区切りのインターフェース名を結合して保持し、構文回帰テストを追加した。 | 2026-07-23 |
| 54 | 実行互換性 #54: `Line Input` の配列要素 | **Bug 54-A 修正済み**: `Line Input #1, lines(i)` が代入先を識別子に限定して構文エラーになった。配列要素を代入先として受理し、2行のファイル読み込み結果を回帰テスト化した。 | 2026-07-23 |
| 53 | 実行互換性 #53: 省略開始位置の `Lock` | **Bug 53-A 修正済み**: `Lock #1, To 100` と対になる `Unlock` が、開始位置の式を必須として `To` で構文エラーになった。開始位置なしの範囲を表現し、開始レコード 1 として競合判定・解除できるよう回帰テスト化した。 | 2026-07-23 |
| 52 | 実行互換性 #52: `Dim Shared` | **Bug 52-A 修正済み**: VBA6 互換のモジュール変数 `Dim Shared x As Long` で `Shared` を変数名として解釈して構文エラーになった。修飾情報を AST に保持して受理し、モジュール変数の状態維持を回帰テスト化した。 | 2026-07-23 |
| 51 | 実行互換性 #51: 代替比較演算子 | **Bug 51-A 修正済み**: `><`、`=<`、`=>` が別々の演算子トークンになり、比較式と `Case Is >< 5` をパースできなかった。字句解析でそれぞれ `<>`、`<=`、`>=` に正規化し、式と `Select Case` の実行結果を回帰テスト化した。 | 2026-07-23 |
| 50 | 実行互換性 #50: `If Then` 行番号分岐 | **Bug 50-A 修正済み**: `If True Then 100` の `100` をラベル宣言として扱い、条件付きジャンプが起きなかった。単行 `If` の数値節を `GoToStatement` として扱い、Then / Else の数値ラベルを回帰テスト化した。 | 2026-07-23 |
| 49 | 実行互換性 #49: 2語形式 `Go Sub` | **Bug 49-A 修正済み**: `Go Sub worker` が `Go` を識別子として扱いパースエラーだった。2語を `GoSubStatement` として受理し、`Return` を経由した名前・数値ラベルの回帰テストを追加した。 | 2026-07-23 |
| 48 | 実行互換性 #48: 2語形式 `Go To` | **Bug 48-A 修正済み**: `Go To done` が `Go` を識別子として扱いパースエラーだった。2語を `GoToStatement` として受理し、名前・数値ラベルへ正しくジャンプする回帰テストを追加した。 | 2026-07-23 |
| 47 | 実行互換性 #47: `With` 内の `Call` | **Bug 47-A 修正済み**: `Call .AddItem "A"` が暗黙 `With` メンバーを呼び出し先として認めずパースエラーだった。括弧なし引数と `ImplicitWithObjectExpression` を `CallStatement` として受理し、クラスメソッドの回帰テストを追加した。 | 2026-07-23 |
| 46 | エンジン評価 #46: `Declare PtrSafe` 構文 | `Private Declare PtrSafe Sub/Function` の `Lib`、`Alias`、`LongPtr`、`ByRef String`、戻り型を含むモジュールを評価。外部 DLL は安全なスタブとして呼ばず、構文解析・ロードと後続のクラス／Collection 処理が正常であることを確認した。 | 2026-07-23 |
| 45 | エンジン評価 #45: パーサーの手続き単位回復 | `errorRecovery: true`（LSP と同じ経路）で、壊れた `Sub` の後にある `StillUsable` を保持しつつ診断を返すことを確認。通常の実行パーサーは例外を送出するが、LSP 用回復経路の仕様どおりでありバグは再現しなかった。 | 2026-07-23 |
| 44 | エンジン評価 #44: 財務関数の反復・配列基底 | **Bug 44-A 修正済み**: `IRR` と `MIRR` が1基底の VBA 配列で未使用の添字0を数値化し `NaN` を返した。**Bug 44-B 修正済み**: `IPmt` は第2期以降の符号が逆で期首払いも誤計算、`per > nper` も受理していた。配列基底・通常／期首払い・範囲外 period を回帰テスト化した。 | 2026-07-23 |
| 43 | エンジン評価 #43: クラス本体の `Dim` / `Static` | クラスフィールドの `Dim`、Collection 初期化・走査、`Static hits As Long` のインスタンス別状態を評価。同一インスタンスは `1:2`、別インスタンスは `1` から開始し、期待どおり `1:2:1` を返した。バグは再現しなかった。 | 2026-07-23 |
| 42 | エンジン評価 #42: 年始の週番号 | `DateAdd("ww")` は年またぎを含め正常。**Bug 42-A 修正済み**: `DatePart("ww")` が `firstweekofyear` を無視し、`#1/1/2021#` で `vbFirstJan1` / `vbFirstFourDays` / `vbFirstFullWeek` のすべてを週1としていた。各規則に応じた週開始日を計算し、結果を `1,53,52` に修正した。 | 2026-07-23 |
| 41 | エンジン評価 #41: Decimal 除算と銀行丸め | `CDec` による除算・符号・28桁丸め・銀行丸め境界を、クラス・Collection・`On Error` を含む集計サンプルで評価。`1/6`、`1/7`、負値、`±1.005` / `±1.015`、`CStr` の不要な末尾ゼロ除去はいずれも期待どおりで、バグは再現しなかった。 | 2026-07-23 |
| 40 | エンジン評価 #40: `Set` 連鎖代入 | クラスの `Property Set`、Dictionary の `Item`、右辺のキー付き取得を組み合わせた配送ルート集計で評価。`Set byCode.Item(key) = obj`、`Set first.NextStop = obj`、`Set byCode.Item("P2").NextStop = byCode.Item("SPARE")` は期待どおり参照を保持し、バグは再現しなかった。 | 2026-07-23 |
| 39 | エンジン評価 #39: GUI 操作スタブ | `AppActivate "Microsoft Excel"` と `SendKeys "{ESC}", True` の文形式を、クラス・Collection・文字列処理・`On Error`・ループを含むサンプルと最小再現で評価。構文・実行とも正常で、GUI へ実際に副作用を与えない `[STUB]` 出力となる。`Wait:=True` も安全な no-op として扱うため、実 GUI 操作・待機はエミュレートしない。 | 2026-07-22 |
| 38 | 実 Excel 差分確認 #38: Double `D` 指数表記 | 実機で確認済みの `1.5D+10` / `2D5` は `Double` として正常に実行できる仕様に合わせた。**Bug 38-A 修正済み**: レキサーが `D`/`d` を指数部として扱わずモジュール全体をパースエラーにしていた。`D`/`d` を受理し、数値変換前に `e` へ正規化した。大文字・小文字・正負指数と `TypeName` を回帰テスト化。 | 2026-07-22 |
| 37 | 実 Excel 差分確認 #37: 文字列から Byte 配列への暗黙変換 | 実機で確認済みの `Dim b() As Byte : b = "Aあ"` を修正。VBA と同じ UTF-16LE コード単位で `41 00 42 30`、下限 0・上限 3 の配列を生成する。 | 2026-07-22 |
| 36 | 実 Excel 差分確認 #36: 日付リテラル・バイナリファイル I/O | 実機で `#3/15#` / `#15/3#` が実行年の 3 月 15 日、時刻併記が保持されることを確認し回帰テスト化。**Bug 26-1/2/4/5 修正済み**: `Long`、`Integer`、CP932 文字列、固定長文字列のみの UDT を `Put` / `Get` でバイナリ入出力できるようにした。実機の `Long 1234567`=`87 D6 12 00`、UDT=連続9バイト、`"Aあ"`=`41 82 A0` を回帰テストに固定。続けて、準備用 `Put #f,,CByte(1)` が型付き変換式を扱えず Error 13 になる回帰、および `Open ... Shared As #f` を構文エラーにするバグを修正。実機5ケースで確定した `Access` / `Lock` 共有行列（Read/Write の禁止と Shared）を実装し、競合時を Error 70 にした。 | 2026-07-22 |
| 35 | エンジン評価 #35: ファイルロック付き在庫台帳 | `Lock` / `Unlock` の範囲指定・解除、閉鎖済みファイル番号、重複範囲を評価。**Bug 35-A 発見・修正済み**: Lock/Unlock が無条件成功するスタブだった。ハンドルごとの範囲ロック管理を追加し、無効ファイル番号・重複範囲・未ロック解除を VBA エラーにした。 | 2026-07-22 |
| 34 | エンジン評価 #34: 年次締切・休業日判定ドメイン | `HolidayCalendar.cls` と標準モジュールで **日付リテラル**・`DateValue`/`TimeValue`・`DateSerial`・`Weekday`・年末日時を評価。Collection、文字列正規化、`On Error GoTo`、ループ、条件分岐を含むサンプルで通常・不正入力を確認。`#3/15#`・`#15/3#`・うるう日・日時併記を調査。**Bug 34-A 発見・修正済み**: 2桁年の日付リテラルが常に 20xx と解釈され、Excel の既定 2029 ルールおよび `DateSerial` と不一致だった。 | 2026-07-19 |
| 33 | エンジン評価 #33: カードゲーム（ブラックジャック風）シミュレーション・ドメイン | **定数からの定数定義・畳み込み**（`DECK_SIZE = SUIT_COUNT * RANK_COUNT`）: 一発動作。**行継続 `_`（引数リスト内・演算子後）/ コロン複文 / REM 行**: 一発動作。**`ChrW`/`AscW` 往復・非 ASCII の `Len`**: 正確（♠=9824）。**With のネスト（`With .TopCard` = Property Get 返却オブジェクト）**: 内側 `.` 解決正常。**`Do While`+`Exit Do` / Randomize シード再現性**: 正常。**Bug 33-A 発見・修正済み**: クラス内の括弧なし自メンバー参照が silent Empty。**Bug 33-B 発見・修正済み**: `Collection.Add` の名前付き `Before:=`/`After:=` が誤バインド。**Bug 33-C 発見・修正済み**: 未使用 `As New` 変数の ByVal 渡しで実体が呼び出し元に反映されない。leniency 記録: `x = 1 REM comment`（コロンなし）を受理（実 VBA は構文エラー、低優先）。 | 2026-07-18 |
| 32 | エンジン評価 #32: 差し込みテンプレートエンジン + 帳票出力ドメイン | **ジャグ配列 `rows(i)(j)`**: クラス内・ローカルとも完全動作。**クラスメソッドから配列返却 + `UBound`/添字**: `ReDim Preserve` 後の返却も正常。**配列の ByRef 渡し**: VBA 内書き戻し・TS `run()` args 経由の書き戻しとも動作。**インデックス付き Property Let/Get `obj.Item(3) = "x"`**: 一発動作。**`Err.Raise vbObjectError+n` の Description 伝播**: 正確。**`Print #` ゾーン**: セミコロン連結・カンマ 14 桁ゾーン・`Spc(n)`・数値書式は正確。**Bug 32-A〜32-E 発見・全件修正済み**: Tab(n) オフバイワン / `#1, #日付#` レクサー誤認 / `Input #` 引用符内カンマ分割 / `Write #` 日付書式・改行コード / `LSet` UDT 間コピー Error 424。仕様準拠確認: `Write #` が埋め込み引用符を二重化しないのは実 VBA と同じ。 | 2026-07-18 |
| 31 | エンジン評価 #31: 再帰下降式パーサー（電卓インタープリター・ドメイン） | **クラス 3 つ + 標準モジュール連携（Token.cls / Tokenizer.cls / ExprParser.cls）**: クラス内から別クラス New・Collection へのオブジェクト格納・`Item(i)`/`For Each` 取り出し 完全動作。**オブジェクト配列**: `Dim tokens() As Token` + `ReDim Preserve` + `Set tokens(i) = New Token` + `For Each` 走査 完全動作。**後置 `Do...Loop While` / `Do...Loop Until`**: 完全動作。**`Err.Raise vbObjectError + n` → `Err.Number - vbObjectError` 往復**: 完全動作（`vbObjectError` = -2147221504 正確）。**Dictionary 暗黙 Let 代入 `m_vars(key) = v`**: 追加・更新とも正常。**文字列連結 20000 回**: 176ms で完走。**行番号付きレガシー構文**: `10 x = 1` 形式のパース・`GoTo 30`・`On Error GoTo 99`（数値ターゲット）すべて動作。**Bug 31-A 発見・修正済み**: `Public/Private Static Sub|Function` がパースエラー。**Bug 31-B 発見・修正済み**: `Erl` が未実装で常に 0（識別子として暗黙解決）。実装追加でハンドラー内 `Erl` がエラー行番号を返すように。 | 2026-07-18 |
| 30 | エンジン評価 #30: 状態機械 / ワークフローエンジン（注文処理ドメイン） | **複数クラス連携（State.cls + WorkflowEngine.cls）**: `Class_Initialize` で `Scripting.Dictionary` をフィールド初期化・複数クラス間の Set 代入・メソッド呼び出し 正常動作。**Collection による履歴記録**: `myCol.Add` / `For Each` による全件走査 正常動作。**On Error GoTo + Err.Raise**: 無効遷移（存在しない状態・イベント）の Error 5 送出・呼び出し元での捕捉 正常動作。**For Each で Dictionary Keys 走査**: `For Each key In dict.Keys` パターン正常動作。**`new VBARunner(optionsObj)` 誤用で ERR_INVALID_ARG_TYPE**: Node.js 内部エラーになりわかりにくい（改善提案）。**Bug 30-A 発見・修正済み**: `VarType(classInstance)` が 36 (vbUserDefinedType) を返す（仕様: 9 = vbObject）。`builtins.ts` の `__vbaClass__` チェックが `__vbaTypeName__` チェックより後にあったため。チェック順を入れ替えて修正。`TypeName()` は正常動作していた。 | 2026-07-17 |
| 29 | エンジン評価 #29: 演算子セマンティクス / Option Base 1・Option Compare Text / クラスイベント（Event/RaiseEvent/WithEvents） / Erase・IsMissing・型サフィックス・While Wend（音楽理論ドメイン: 移調・スケール・コード進行・メトロノーム） | **`\`/`Mod`**: 事前銀行家丸め・負数・ゼロ除算 12 ケース全て実 VBA と一致。**`&` vs `+`**: `"1"+1`=2、`"a"+1`=Error 13、`"a" & Null`="a" 正常（`Null & Null` のみ Bug 29-D）。**比較型強制**: `"10">9`・Empty 比較 全一致。**`Xor`/`Eqv`/`Imp`**: ビット演算として正確（`5 Eqv 3`=-7、`5 Imp 3`=-5）。**クラスイベント**: `Public Event`/`RaiseEvent`/`Dim WithEvents`/`Private Sub obj_Event()` ハンドラー **完全動作**（4拍子6拍のイベント列正確）。**Option Base 1**: `Dim a(3)` の LBound=1 正常（`Array()` のみ Bug 29-E）。**Option Compare Text**: `=`/`<`/`Like`/`InStr` のモジュールスコープ既定まで正常。**Erase**: 動的配列解放後 UBound=Error 9 正常。**IsMissing**: 正常。**サフィックスリテラル**（`123&`/`1.5#`/`1.5!`/`123@`）: TypeName 全正確（サフィックス付き `Dim n&` は Bug 29-A）。**While Wend**: 正常。**Bug 29-A〜29-H 発見**（下記セクション参照・全件メイン再現確認済み） | 2026-07-17 |
| 28 | エンジン評価 #28: 再帰アルゴリズム / UDT配列 / ReDim Preserve / StrConv日本語変換 / Nz / RGB・QBColor / Format Yes/No（アルゴリズムライブラリ+テキスト分析ドメイン） | **再帰 Function（Fibonacci・MergeSort・BST 走査）**: 完全正常動作。スタック・スコープに問題なし。**StrConv 日本語変換**: `vbKatakana(16)` / `vbHiragana(32)` / `vbWide(4)` / `vbNarrow(8)` すべて期待通り動作。**Nz**: `Nz(Null, 0)` / `Nz(Empty, 42)` / `Nz(有値, default)` の3パターン正常。**RGB / QBColor**: 評価 #27 修正済みで正常動作を確認。**Format("Yes/No"/"On/Off"/"True/False")**: 評価 #27 修正済みで正常動作を確認。**TextAnalyzer.cls**: Class_Initialize・On Error GoTo・Property Get・Scripting.Dictionary 組み合わせ正常動作。**Bug 28-1 発見**: `ReDim Preserve` で UDT 配列を拡張した後、新インデックス要素のメンバーアクセスが Error 424。 | 2026-07-16 |
| 28 | エンジン評価 #28: 再帰アルゴリズム / UDT配列 / StrConv日本語変換 / Nz・RGB・Format（アルゴリズムライブラリ + テキスト分析ドメイン） | **再帰 Function（Fibonacci・MergeSort・BST走査）**: 完全正常動作。スタック・スコープに問題なし。**StrConv 日本語変換（vbKatakana=16/vbHiragana=32/vbWide=4/vbNarrow=8）**: すべて期待通り動作。**Nz / RGB / QBColor / Format "Yes/No"**: 評価直前のバグ修正（Bug A〜N）によりすべて正常動作。**TextAnalyzer.cls（Class_Initialize + Dictionary + On Error GoTo + Property Get）**: 完全正常動作。**Bug 28-1 発見・修正済み**: `ReDim Preserve n(0 To 1)` で UDT 配列を拡張後、`n(1).Value = 2` が Error 424。旧 `fillArrayWithUDT` の `!isPreserve` ガードにより新インデックスが未初期化のまま残っていた。 | 2026-07-16 |
| 27 | エンジン評価 #27: Registry スタブ / 時刻関数 / Filter・Array・CallByName / Shell・DoEvents（人事・勤怠管理システムドメイン） | **SaveSetting/GetSetting/DeleteSetting/GetAllSettings**: すべて正常動作。仮想レジストリに設定を保存・取得・削除・全取得できる。`GetAllSettings` は `Array[row][0=key, 1=val]` の2次元配列を返す。**TimeSerial/TimeValue/Hour/Minute/Second**: すべて正常動作。`TimeSerial(8,30,0)` → `Hour()=8, Minute()=30, Second()=0` 正確。`TimeValue("16:20:05")` も正常。`#HH:MM:SS#` 時刻リテラル（修正済み）も動作。日付+時刻リテラル `#2024/01/15 08:30:00#` も正常。**Filter()**: 正常動作。`Filter(arr, keyword)` デフォルト大文字小文字区別あり。`vbTextCompare(1)` で区別なし。`Filter(arr, keyword, False)` 非一致フィルタも正常。一致なしで空配列（`UBound=-1`）も正常。**Array() 関数**: `Dim v As Variant : v = Array(1, "x", 3.14)` で混合型配列格納が正常動作。**CallByName**: `VbGet(2)`（プロパティ GET）・`VbMethod(1)`（メソッド呼び出し）は正常動作。`VbLet(4)`（プロパティ SET）は **Bug 27-1** として Error 438 が発生（未実装）。**Shell**: `Shell("notepad.exe")` → `[SHELL]` ログを出力しタスク ID `1` を返すスタブ。正常動作。**DoEvents**: `DoEvents` → `0` を返す no-op。正常動作。**Bug 27-2 発見**: `On Error GoTo label` ハンドラー内で `Exit Function/Sub` で抜けると呼び出し元の `Err.Number` に エラー番号が残留する。`Resume` 経由の場合のみ `errObj.clear()` が呼ばれるため、`Exit Function` での正常リターン後も `Err.Number` がリセットされない。 | 2026-07-15 |
| 26 | エンジン評価 #26: バイナリファイル I/O / ランダムアクセスファイル / ファイルシステム関数 / Error(n)（バイナリファイル処理ドメイン） | **Open For Binary**: オープン・Close 自体は正常。`Put #n,,var`/`Get #n,,var` の実装が完全テキストベース（String(data)→UTF-8→書き込み、1024バイト一括読み込み）のため、型サイズに基づくバイナリ I/O が機能しない（**Bug 26-1/26-2**）。**Open For Random Len=**: `Len = recLen` 節が parser に未実装→Parse error（**Bug 26-3**）。**UDT Put/Get**: `String(udtObj)="[object Object]"` を書き込む / `Get` は UDT に文字列を代入しようとして Error 424（**Bug 26-4/26-5**）。**FileLen**: 正常動作。**FileDateTime**: 正常動作（`fs.statSync().mtime` から VbaDate を生成）。**Kill**: ステートメント形式で正常動作。**GetAttr**: Error 35 未実装（`setattr` は登録済みだが `getattr` 未登録）（**Bug 26-7**）。**Error(n)**: 主要エラーコード（5/6/9/11/13/53/91）は正しいメッセージを返す。未登録番号（7=Out of memory、14、0、999）はすべて "Application-defined or object-defined error" にフォールバック。 | 2026-07-14 |
| 25 | エンジン評価 #25: Dir() / DatePart・MonthName・WeekdayName / Byte関数（ログファイル解析ドメイン） | **Dir()**: `Dir("C:\logs\*.log")` → VFS 上の `.log` ファイル3件を列挙。引数なし繰り返し呼び出し（`Do While f <> ""`）パターン正常動作。存在しないパターンは `""` を返す（エラーなし）。ネスト Dir() はグローバル状態リセットで元のポインターが失われる（実 VBA 仕様準拠）。**DatePart**: `DatePart("yyyy"/"m"/"d"/"ww"/"h"/"n"/"s"/"q", d)` 全インターバル正常動作。ただし第3引数（firstdayofweek）は Error 450（未対応）。**MonthName**: `MonthName(3)="March"` / `MonthName(3, True)="Mar"` 正常。**WeekdayName**: `WeekdayName(1)="Sunday"` / `WeekdayName(1, True)="Sun"` / 第3引数も正常動作（DatePart との非対称あり）。**MidB**: `MidB("Hello", 1, 2)="H"` 正常（UTF-16LE で 1文字=2バイト仕様通り）。**Bug #25-1〜#25-4 発見**（未修正）: `LenB` / `AscB` / `ChrB` が Error 35 未実装。`Dim b() As Byte : b = str` が Error 424 未実装。**Bug #25-5 発見**（未修正）: `Split(str, delim, limit)` 第3引数が Error 450。**Bug #25-6 発見**（未修正）: `DatePart` 第3引数（firstdayofweek）が Error 450。 | 2026-07-13 |
| 24 | エンジン評価 #24: 財務関数 / Mid代入 / LSet・RSet（住宅ローン計算・投資シミュレーションドメイン） | **財務関数**: `Pmt`（99,379円≒期待値99,378円）/ `IPmt`（利息50,000円）/ `PPmt`（元金49,379円、`IPmt+PPmt=Pmt` 確認）/ `NPer`（323.7ヶ月）/ `FV`（139万円≒期待値139万円）/ `SLN`（1,800円）/ `SYD`（year1=3,000円、year5=600円）/ `DDB`（year1=4,000円、year2=2,400円）/ `Rate`（収束にはguessRateが重要、初期値が遠いと別の根に収束する点は仕様通り）— 全関数が実装済みで正常動作。**Mid 代入形式（`Mid(s, i, n) = val`）**: 正常動作。`Mid(s, 7, 5) = "VBA!!"` → "Hello VBA!!" / `Mid(s, 2, 2) = "XY"` → "AXYDE" / `eval()` 直接呼び出しも正常。**LSet / RSet**: 固定長文字列に対して正常動作。`LSet fixedStr = "ABC"` → "ABC       "（10文字）/ `RSet fixedStr2 = "XYZ"` → "       XYZ"（10文字）。`FormatReportLine` 関数での LSet（ラベル）+ RSet（金額右詰め）パターンも正常動作。**Collection + string処理**: `SummarizeLoans` でコレクション + `Pmt` 関数の組み合わせが正常動作。**Bug 24-1 発見・修正済み**: `NPV(rate, flows)` に VBA 1-based 配列（`Dim flows(1 To 3)`）を渡すと NaN を返す。`builtins.ts` の NPV 実装が `values.map(Number)` でインデックス 0 から反復するため、`vbaBase=1` の配列ではインデックス 0（undefined）が NaN になっていた。`(values as any).vbaBase ?? 0` で基底インデックスを取得し正しく反復するよう修正。レグレッションテスト: `tests/spec/financial.test.ts` Bug 24-1 ブロック。 | 2026-07-13 |
| 23 | エンジン評価 #23: 多次元配列 / Currency・Decimal BigInt 精度（売上管理・収益分析ドメイン） | **テーマA（多次元配列）**: `Dim arr(3,4) As Currency` 2D配列宣言正常。`UBound(arr,1)` / `UBound(arr,2)` 各次元の上限取得正常。`ReDim Preserve` による最終次元（2次元目）拡張正常・既存データ保持確認。クラスメンバーとしての 2D 配列（`SalesMatrix.cls`）正常動作。配列を Function 戻り値として返す（`GetDataCopy`）正常動作。**テーマB（Currency/Decimal 精度）**: `CCur(0.1) + CCur(0.2) = "0.3"`（CStr で厳密一致）確認。`CDec(1)/CDec(3) = "0.3333333333333333333333333333"`（28桁）確認。`TypeName(CDec(x)) = "Decimal"` / `VarType(CDec(x)) = 14`（vbDecimal）正常。`Err.Raise` カスタムメッセージが JS 側に正しく伝わる。**Bug C-1 発見・修正済み**: `run()` / `eval()` が `VbaCurrency` を JS `number` に変換せず `VbaCurrency { internal: BigInt }` オブジェクトのまま返す。`test-runner.ts` に `normalizeVbaValue` ヘルパーを追加し `VbaCurrency` / `VbaDecimal` を `Number(v.toString())` で変換するよう修正。**Bug C-2 発見・修正済み**: Currency 型 ByRef パラメーターへの書き戻し後に `JSON.stringify(BigInt)` でクラッシュ。`formatVbaArg` ヘルパーで `VbaCurrency` / `VbaDecimal` を `JSON.stringify` 経由せずに文字列化するよう修正。レグレッションテスト: `tests/test-libs-tests/vba-currency-decimal-normalize.test.ts`（6テスト）。 | 2026-07-13 |
| 22 | エンジン評価 #22: 数学・数値関数 / 型チェック / Like / Optional+ParamArray / 名前付き引数 / IIf・Choose・Switch / Null / 文字列変換（暗号・エンコーディング+数値解析ドメイン） | **Hex/Oct**: `Hex(65)="41"`, `Oct(65)="101"` 全正常。**Like演算子**: `"*@*.*"` メールバリデーション / `?` 単一文字 / `[a-c]` 文字クラス / `[!acd]` 否定クラス / `#` 数字 すべて正常動作。**StrReverse/Space/String**: `StrReverse("Hello")="olleH"` / `Space(3)="   "` / `String(5,"-")="-----"` 正常。**数学関数**: `Sqr`, `Abs`, `Log`, `Exp`, `Sin`, `Cos`, `Atn`（Pi計算: `4*Atn(1)`）, `Sgn`, `Fix` すべて正常。Fix(-2.7)=-2, Int(-2.7)=-3 の違いも正確。**ParamArray**: `CalcMean(1,2,3,4,5)=3`, `CalcStdDev(2,4,4,4,5,5,7,9)≈2.138` 正常。**Rnd/Randomize**: `Randomize seed` でシード固定 → 同一シードで同一乱数列を確認。**TypeName/VarType**: `TypeName(42)="Integer"`, `TypeName(32768)="Long"`, `TypeName(3.14)="Double"` リテラル引数は正常。VarType constants (0=Empty, 1=Null, 2=Integer, 3=Long, 5=Double, 8=String, 11=Boolean) も正確。IsNull/IsEmpty/IsArray/IsObject/IsError/CVErr すべて正常。**Null**: `Null` リテラル代入 / `IsNull(Null)=True` / `IsEmpty(Null)=False` / `Null+1=Null`（伝播）/ `TypeName(Null)="Null"` すべて正常。**IIf/Choose/Switch**: 全正常。`Choose(1,"Mon","Tue",...)`, `Switch(score>=90,"A", ...)` も正常。**Optional/Named Args**: `FormatNum(value:=3.14159, decimals:=3)` 名前付き引数・引数順序逆転・部分指定のみ すべて正常（今回初確認）。**StrConv**: `StrConv("hello",1)="HELLO"` / `StrConv("HELLO",2)="hello"` 正常。**Bug 22-1 発見・修正済み**: `Dim v As Variant : v = 42 : TypeName(v)` → "Double"（正しくは "Integer"）。`evaluateTypeIntrinsic` の非リテラル数値分岐に `Number.isInteger(val)` + Integer/Long 範囲チェックを追加して修正。レグレッションテスト: `tests/spec/typename.test.ts` Bug 22-1 ブロック。 | 2026-07-12 |
| 21 | エンジン評価 #21: Static 変数 / Single 型 / 固定長文字列 / Select Case 範囲 / Do Until（シリアルナンバー生成システムドメイン） | **Static 変数**: 複数回の `run()` をまたいで `seqCounter` が正確にインクリメント。完全正常動作。**Single 型・型変換**: `CSng(1)/CSng(3)` → `0.3333333432674408`（32bit 精度）、`CDbl(1)/CDbl(3)` → `0.3333333333333333`（64bit 精度）。実 VBA と同様の精度差を正しく再現。**Select Case 範囲パターン**: `Case Is <= 10` / `Case 11 To 30` / `Case 56 To 80` / `Case Is >= 100` の全パターン正常動作。**Do Until / Loop Until**: `Do Until cond` はゼロ回実行・`Do ... Loop Until cond` は条件成立時でも最低1回実行を確認。両形式正常動作。**固定長文字列 `Dim s As String * N`**: **Bug 21-1 発見**（未修正）: パーサーが `*` をステートメント後の予期しないトークンとして扱い Parse error。`VariableDeclarator` AST インターフェース・`parseDimStatement`・エバリュエーターのいずれも `* N` に未対応。完全未実装。 | 2026-07-11 |
| 20 | VS Code 拡張機能 評価 #20: UDT/Enum/GoSub/Implements エンジン + LSP 統合（地図座標管理ドメイン） | **テーマA（UDT/Enum）**: UDT フィールドアクセス・ByRef 引数渡し・戻り値返却・固定長配列・ネスト UDT（`ln.Start.X`）・Enum 定数（修飾あり/なし）・Select Case での Enum 使用 全正常動作。**テーマB（GoSub/Return）**: 基本実行・複数回呼び出し・複数ラベルシーケンス・ネスト GoSub・`Return without GoSub` エラー 全正常動作。戻りアドレスのスタック管理が正しく動作。**テーマC（Implements）**: 構文エラーなし・直接呼び出し・インターフェース経由ディスパッチ・ポリモーフィズム（Circle/Rectangle 切り替え）全正常動作。**テーマD（LSP 統合）**: 診断（UDT/Enum で誤 VBA016 なし）正常。補完/ホバー/フォーマット に Bug LSP-1/2/3 発見・修正済み。フォーマッターの `=` 前後スペース未整形は設計上の制限（Bug LSP-4）。 | 2026-07-07 |
| 19 | VS Code 拡張機能 評価 #19: buildExtractFunctionEdit edge case / Dead Store 精度 / analyzeDefUse 複雑制御フロー（学力テスト結果管理ドメイン） | **テーマA（buildExtractFunctionEdit edge case）**: 複数出力パラメーター正常（ByRef x/y/z 生成）・純粋ローカル変数のみ選択で空引数呼び出し正常・Function 内抽出は常に Sub 生成（型推論不可）。**パラメーター名小文字化**: analyzeDefUse が返す変数名が小文字正規化済みのためシグネチャの `inputVal` → `inputval` になる（改善候補）。**テーマB（Dead Store 精度）**: ループカウンター変数・集計変数・条件分岐内代入・関数戻り値変数 すべて誤検出なし ✓ 実際のデッドストア正しく検出 ✓ CFG ベース生変数解析が高精度。**テーマC（analyzeDefUse 複雑制御フロー）**: With ブロック内変数 inputs/outputs/locals 正しく分類 ✓ ByRef 引数渡し後の後続使用を outputs に正しく分類 ✓ **テーマD（コード品質）**: 全パラメーターが As Variant（型推論なし）は設計制限。選択外 Dim の extraDims も As Variant になる。**Bug W1 発見・修正済み**: With ブロック内でユーザー定義クラスの Property Get（`.Score`）が Error 424 になる。Function 呼び出し（`.Summary()`）は正常。`evaluateImplicitWithObjectExpression`（`evaluator.ts:6268`）に `__vbaClass__` ブランチが欠如。`evaluateMemberExpression` と同じ Property Get → callClassMethod パターンを追加して修正。レグレッションテスト: `class-module.test.ts` 末尾 Bug W1 ブロック。 | 2026-07-07 |
| 18 | VS Code 拡張機能 評価 #18: LSP リファクタリングプロバイダー群 | **`analyzeDefUse(proc, startLine, endLine)`**（`src/engine/def-use-analyzer.ts`）: inputs/outputs/locals 分類正常。ループ・条件分岐ケースも正確。**`findDeadStores(proc)`**（`src/engine/dead-store.ts`）: デッドストア検出正常。ByRef パラメーター除外・関数戻り値変数の alwaysLive 扱い正しい。**`LspServer.getCodeActions(uri, range)`**: range 正規化（ドラッグ選択の end.character===0）・プロシージャ外選択で [] を返す 正常動作。**`LspServer.buildExtractFunctionEdit(...)`**: **Bug R1 発見・修正済み**: 選択範囲に `Dim x` が含まれ `x` が ByRef 出力パラメーターになる場合、`Dim x` と `ByRef x As Variant` が共存して VBA コンパイルエラーになる不正コードを生成していた。`reindented` 生成時に inputs/outputs 変数の `Dim` 行をフィルタリングする修正を適用。レグレッションテスト追加（`lsp-code-actions.test.ts` Test 13）。**未実装機能**: Introduce Variable / Extract Constant / Inline Variable / Introduce With / Organize Declarations は `src/lsp/` に存在しない。 | 2026-07-07 |
| 17 | VS Code 拡張機能 評価 #17: InlayHint / Diagnostics / Completion / Hover 全面評価 | **InlayHint（inferVariantTypes/inferProcedureHints）**: A1-A7 全正常。CreateObject型推論・曖昧型ヒントなし・数値代入→Long・型なしパラメーター→Variant・モジュールレベルObject推論・戻り型ヒント 全動作確認。`Dim x As Variant`+数値代入→Long ヒントが出る（設計通り）/ **CompletionProvider**: ユーザー定義クラスチェーン(SimSheet→Cells→SimCell.Value)・CreateObject Dictionary補完・型不明Object/Variant空補完 全正常 / **DiagnosticProvider（lintProgram+checkUnknownTypes）**: VBA001〜VBA014 全lint コード確認済み（VBA011除く・Excel依存のみ）。severity: VBA013=Error(1)・VBA003=Hint(4)・VBA002/VBA004=Info(3)・他Warning(2) / **HoverProvider**: ローカル変数・パラメーター・モジュールレベル変数・定数・プロシージャ・As Object→CreateObject型推論表示・到達定義情報 全正常動作 / **Bug E1**: `Const MAX As Long = 100` のホバーが `Const MAX` のみ（型・値を表示しない）`symbol-table.ts:addConstDeclaration` の `displayText` が `Const 名前` のみ / **Bug E2**: `Public`/`Private` キーワードが hover で lowercase 表示（`public count As Long`）パーサーが scope を lowercase 格納するため / VBA003(ByRef/ByVal省略)・VBA012(ByRef明示なし代入)も確認済み | 2026-07-07 |
| 16 | VS Code 拡張機能 評価 #15 バグ修正確認 + 新規 edge case | **全3件修正確認**: `autoParensEdit` 戻り型付き Function 対応済み / `TestRunner.runTests(src)` 本実装済み（`Err.Raise` → `failed` を返す）/ `runTestWithEvaluation` エラーメッセージ正常化済み / **新規**: VBA キーワード（`Set`/`If`）参照は正しく空を返す（問題なし）/ `Dim x As`（型名欠落）はパーサーが黙って回復し構文エラーを出さない（改善候補）/ `ByRef/ByVal` 省略への VBA003 警告が severity:Warning（新規ユーザーには noisy） | 2026-07-05 |
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
| 10 | VS Code 拡張 LSP 機能（評価 #9 修正確認 + 新規テーマ） | 引数付きチェーン補完の修正確認（OK）/ VBA016 波下線位置の修正確認（OK）/ With ブロック内引数付きチェーン補完（新バグ）/ クロスモジュール補完（`parseAsClass` 必須と判明）/ `generateDefaultTypeStubsJson` JSON valid 性確認 / `setTypeStubs(Map)` API 非対称設計の発見 / カスタム型上書き優先（正常動作） | 2026-07-04 |
| 11 | VS Code 拡張 LSP 機能（評価 #10 修正確認 + シグネチャヘルプ + Quick Fix） | With ブロック内引数付きチェーン補完修正確認（OK・48 件）/ `parseTypeStubsJson` API 非対称解消確認（OK）/ `SignatureHelpProvider` 組み込み 60 件超・ユーザー定義 Function・ネスト呼び出し（内側優先）正常動作 / **文字列リテラル入力中にシグネチャが消えるバグ発見 → Lexer ベース修正済み** / VBA016 Quick Fix 実装確認（`extension.ts` に `initTypeStubs`・`addToTypeStubs` コマンド実装済み）/ `parse()` は `Program` を返す（`getCompletions` には `.body` を渡す必要あり） | 2026-07-04 |
| 12 | `__mocks__` 注入・`setBuiltinOverride`・FileSystemWatcher（評価 #12） | JS `__addCreateObject__` による `CreateObject` 置換（Excel.Application）/ VBA `.cls` クラスモック差し込み（Logger.cls）/ `__mocks__.bas` 単一ファイル形式 / `setBuiltinOverride` で MsgBox 戻り値固定（vbOK/vbCancel 再現確認済み）/ `spy()` API 呼び出し引数記録 / 配列渡し時モックスキャンなし確認 / `.ts` モック動作確認 / `__progId__` なし factory のキー補完 / `__mocks__/Class.cls` が本番 `.cls` を上書き（正常）/ **バグ発見・再現確認済み: 複数 VBA `.bas` モックが同名 Public 関数を持つと Ambiguous procedure エラー** / FileSystemWatcher 実装確認（`extension.ts:70-77`）/ `onDidDelete` 未実装 | 2026-07-04 |
| 13 | VS Code 拡張 LSP ナビゲーション（definition / references / rename / symbol）（評価 #13） | `DefinitionProvider.getDefinition` / `ReferencesProvider.getReferences` / `RenameProvider.getRename` / `SymbolProvider.extractSymbols` / `.bas` 全機能正常動作 / セクションヘッダー `' --- Name ---` → SymbolKind.Namespace 抽出 / ローカル変数スコープ絞り込み正常 / クロスモジュールは `null` 返却（設計通り）/ **Bug B 発見・再現確認・修正済み: `parseAsClass` で生成した `ClassDeclaration`/`ProcedureDeclaration` の `loc` が `undefined`（`parseClassBody` が `parseStatement` を経由せず直接パース関数を呼ぶため）→ `parseClassBody` 内で各文パース後に `loc` を設定するよう修正。`tsc -b` / `class-module.test.ts` 全通過確認済み** / Bug A（仕様準拠）: `FuncName = value` の戻り値代入行が refs に含まれる（正規表現全出現検索の仕様。Rename では正しい動作） | 2026-07-05 |
| 14 | VS Code 拡張 LSP（Formatter / CodeLens / FoldingRange / CallGraph）（評価 #14） | `format(source, opts)` + `applyEdits` / キーワード大文字小文字正規化 / ブロックインデント正規化 / Select Case・ElseIf・With・GoTo ラベル・行継続 / ユーザー定義識別子ケーシング / `CodeLensProvider.getCodeLens` + `getDeadCodeWarnings` / `Test_*` 検出・`isEventHandler` / テスト結果注入 / `FoldingRangeProvider.getFoldingRanges` / 全ブロック型対応・単行 If 除外 / `CallGraphProvider.buildCallGraph` / マルチファイル / 相互再帰・自己再帰 / Excel 依存検出 / **Bug: `Test_*` プロシージャが常に `✓ Tested` を返す（疑似陽性）** / 設計ギャップ: 組み込み型名（`String`/`Integer` 等）と組み込みオブジェクト（`Debug`/`Err` 等）は大文字小文字を正規化しない | 2026-07-05 |
| 15 | VS Code 拡張 LSP（auto-parens / keyword-casing / line-continuation / label-navigator / ast-comparison / test-discovery / test-runner / hover-provider / variant-type-inferencer）（評価 #15） | `autoParensEdit`/`getBlockEnd`/`needsBodyIndent`/`needsEndBlock` 各動作確認 / `canonicalKeyword`/`isInStringOrComment` 正常動作 / `needsLineContinuation`/`stripInlineComment` 正常動作 / `LabelNavigator` GoTo定義・参照両方向ナビゲーション正常 / `astEqual`/`serializeAst`/`findMatchingExpressions` 正常動作 / `TestDiscovery.discoverTests` 正常動作 / `TestRunner.runTests` スタブ（常に passed）バグ / `TestRunner.runTestWithEvaluation` エラーメッセージ `"[object Object]"` バグ / `HoverProvider.getHoverInfo` 変数宣言・パラメーター・到達定義情報 正常動作（パラメーター range は col 0 固定・設計上の制限）/ `inferVariantTypes` String/Long/Double/Boolean/ProgID推論 正常動作 / **Bug: `autoParensEdit` が戻り型付き Function（`Function GetValue As Long`）を検出しない** / **Bug: `TestRunner.runTests()` スタブ実装・常に passed** / **Bug: `runTestWithEvaluation` の catch で VBA エラーオブジェクトが `"[object Object]"` になる** | 2026-07-05 |

---

## 発見した問題と対応状況

### 修正済みバグ

| 問題 | 最小再現コード | 修正コミット |
|---|---|---|
| `eval()` で組み込み関数戻り値への `+`/`-` 演算が Error 424 | `r.eval('UBound(arr) + 1')` → Error 424（括弧ワークアラウンド: `(UBound(arr)) + 1`）| `ec63519` |
| `run()` ログで JS 配列引数が `[Object]` と表示される | `r.run('Proc', [[1,2,3]])` → ログが `Proc([Object])` | `ec63519` |
| `Dictionary.Item("nonexistent")` がキーを自動生成しない | 実 VBA では存在しないキーへの `.Item` 読み取りで Empty のエントリを自動生成する（Count+1, Exists→True）。修正後は VBA 互換動作＋コンソール警告を出力 | `ca409b7` |
| `Write #` で Boolean が `#TRUE#`/`#FALSE#` でなく `True`/`False` になる | `evaluateWriteStatement` に `VbaBoolean` 分岐を追加 | `9e25adc` |
| **Bug 22-1: `Dim v As Variant : v = 42 : TypeName(v)` → "Double"（正しくは "Integer"）** | `evaluator.ts:evaluateTypeIntrinsic` の非リテラル数値分岐に `Number.isInteger(val)` + Integer/Long 範囲チェックを追加。`tests/spec/typename.test.ts` Bug 22-1 ブロックにレグレッションテスト追加 | 評価 #22 |
| ~~**Bug #25-1〜3: `LenB` / `AscB` / `ChrB` 未実装**~~ | **修正済み**: `builtins.ts` に `lenb`（文字数×2）/`ascb`（先頭バイト & 0xFF）/`chrb`（charCode & 0xFF → 文字）を登録。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug #25-1〜3 ブロック。 | |
| ~~**Bug #25-4: `Dim b() As Byte : b = str` 未実装**~~ | **評価 #37 で修正済み**: UTF-16LE のコード単位へ展開し、`"Aあ"` → `41 00 42 30` を返す。 | |
| ~~**Bug #25-5: `Split(str, delim, limit)` 第3引数未対応**~~ | **修正済み**: `builtins.ts` の `split` 登録に `limit`（省略可、デフォルト -1）と `compare`（省略可）を追加。`limit=0` は空配列、`limit<0` は全分割、`limit>0` は最大要素数（最後の要素に残り全体）。レグレッションテスト: `tests/spec/split-join.test.ts` Bug #25-5 ブロック。 | |
| ~~**Bug #25-6: `DatePart` 第3引数（firstdayofweek）未対応**~~ | **修正済み**: `builtins.ts` の `datepart` 登録に `firstdayofweek`（省略可、デフォルト 1=日曜）と `firstweekofyear`（省略可）を追加。`'w'` は `(weekday - weekStart + 7) % 7 + 1`、`'ww'` は Jan1 からの日数オフセットで週番号計算。レグレッションテスト: `tests/spec/datetime.test.ts` Bug #25-6 ブロック。 | |
| **Bug 24-1: `NPV(rate, flows)` に VBA 1-based 配列を渡すと NaN** | `Dim flows(1 To 3) As Double : flows(1)=30000 : ... : NPV(0.1, flows)` → NaN | `builtins.ts` の NPV が `values.map(Number)` でインデックス 0 から反復するため `vbaBase=1` の配列でインデックス 0（undefined）が NaN になった。`(values as any).vbaBase ?? 0` で基底インデックスを取得して修正。レグレッションテスト: `tests/spec/financial.test.ts` Bug 24-1 ブロック | 評価 #24 |
| **Bug C-1: `run()` / `eval()` が `VbaCurrency` を `number` に変換しない** | `test-runner.ts` に `normalizeVbaValue` ヘルパーを追加し、`VbaCurrency` は `Number(v.toString())`、`VbaDecimal` は `Number(v.toString())` で変換。`run()` と `eval()` の両メソッドで適用。レグレッションテスト: `tests/test-libs-tests/vba-currency-decimal-normalize.test.ts` | 評価 #23 |
| **Bug C-2: Currency 型 ByRef パラメーター書き戻し後に `JSON.stringify(BigInt)` でクラッシュ** | `test-runner.ts` に `formatVbaArg` ヘルパーを追加し、`VbaCurrency` / `VbaDecimal` を `JSON.stringify` 経由せずに文字列化。`run()` の `formatArgs` 生成で使用。レグレッションテスト: `tests/test-libs-tests/vba-currency-decimal-normalize.test.ts` | 評価 #23 |
| **Bug #25-1〜3: `LenB`/`AscB`/`ChrB` 未実装** | `builtins.ts` に UTF-16LE バイトモデルで実装（`LenB`=文字数×2, `AscB`=先頭バイト, `ChrB`=バイト→文字）。レグレッションテスト: `tests/spec/builtin-strings.test.ts` | 評価 #25 修正 |
| **Bug #25-5: `Split` の `limit` 引数未対応** | `builtins.ts` の split 登録に `limit`/`compare` を追加。VBA 仕様通り（limit=0→空配列、limit>0→最大N要素）。レグレッションテスト: `tests/spec/split-join.test.ts` | 評価 #25 修正 |
| **Bug #25-6: `DatePart` の `firstdayofweek` 引数未対応** | `builtins.ts` の datepart に `firstdayofweek`/`firstweekofyear` を追加。`'w'`/`'ww'` の曜日計算で weekStart オフセットを反映。レグレッションテスト: `tests/spec/datetime.test.ts` | 評価 #25 修正 |
| **Bug 26-3: `Open For Random Len=N` パースエラー** | `parser.ts:parseOpenStatement` でファイル番号後に `Len = <expr>` をオプション消費するよう修正（`peek().type===Identifier && value.toLowerCase()==='len'`）。レグレッションテスト: `tests/spec/filesystem.test.ts` | 評価 #26 修正 |
| **Bug 26-7: `GetAttr` 未実装 + ファイル属性定数未登録** | `evaluator.ts` に `getattr`（0 返却）/`setattr`（no-op）を登録。`builtins.ts:registerConstants` に `vbNormal`〜`vbAlias` 8定数を追加。レグレッションテスト: `tests/spec/filesystem.test.ts` | 評価 #26 修正 |
| **Bug 26-1/2/4/5: `Put` / `Get` が型サイズを無視し、UDT を文字列化する** | `Long 1234567` を `87 D6 12 00`、`Long`/`Integer`/`String * 3` の UDT を9バイト連続、`"Aあ"` を CP932 の `41 82 A0` として入出力する。実機結果を `tests/spec/binary-file-io.test.ts` に固定。 | 評価 #36 修正 |
| FSO `TextStream.ReadAll()` が `ReadLine()` 後も全体を返す | `readall` が `pos` を参照するよう修正 | `9e25adc` |
| **Bug F: `Format(True, "0")` が "-1" でなく "True" を返す** | `formatFunc` の非 named フォーマット分岐で `VbaBoolean` を `val.value`（数値）に unwrap してから `formatNumber` に渡すよう修正。`Format(True, "0")` → "-1"、`Format(False, "0")` → "0" が正しく返るようになった。レグレッションテスト: `tests/spec/builtins.test.ts` Bug F ブロック。 | `8562e4f` |
| **Bug G: `IsDate(1)` が False を返す（数値シリアル日付を認識しない）** | `isdate` 関数に `typeof val === 'number' && isFinite(val)` → `vbaTrue` 分岐を追加。VBA では数値は日付シリアルとして有効な日付。レグレッションテスト: `tests/spec/builtins.test.ts` Bug G ブロック。 | `8562e4f` |
| **Bug H: `Asc("")` が null/NaN を返す（Error 5 にならない）** | `ascFunc` に `str.length === 0` チェックを追加して Error 5 を投げるよう修正。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug H ブロック。 | `8562e4f` |
| **Bug I: `Left/Right` の負値長さ、`Mid` の start<1 や負値長さが Error 5 にならない** | `leftFunc`/`rightFunc` に `l < 0` チェック、`midFunc` に `st < 1` と `len < 0` チェックを追加して Error 5 を投げるよう修正。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug I ブロック。 | `8562e4f` |
| **Bug J: `Space(-1)` / `String(-1, "x")` が JS RangeError（VBA Error 5 にならない）** | `spaceFunc`/`stringFunc` に `count < 0` チェックを追加して VBA Error 5 を投げるよう修正。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug J ブロック。 | `8562e4f` |
| **Bug K: `InStr(0, "abc", "b")` が Error 5 にならずに検索成功する** | `instrFunc` に `Number(start) < 1` チェックを追加して Error 5 を投げるよう修正（2引数形式（start なし）は start=1 扱いのため影響なし）。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug K ブロック。 | `8562e4f` |
| **Bug L: `Format(CCur/CDec, 数値パターン)` が書式未適用で文字列化される** | `formatFunc` の namedFormats 分岐と非 named 分岐の両方で `VbaCurrency`/`VbaDecimal` を `Number(val.toString())` に変換してから `formatNumber` に渡すよう修正。`Format(CCur(1234.5), "#,##0.00")` → "1,234.50" が正しく返る。レグレッションテスト: `tests/spec/builtins.test.ts` Bug L ブロック。 | `5f7abcb` |
| **Bug M: `Chr(256)` が Error 5 にならず Unicode 文字を返す** | `Chr`/`Chr$` は ANSI 範囲 0-255 のみ受け付ける VBA 仕様に従い、範囲外で Error 5 を投げるよう修正。`ChrW`/`ChrW$` は別実装に分離し 0-65535 範囲を適用。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug M ブロック。 | `a611127` |
| **Bug N2: `Format(n, "Scientific")` が小文字 e・1桁指数を返す** | `format.ts` で `toExponential(2)` の出力を VBA 仕様の大文字 E・最低2桁指数（"1.23E+06"）に変換するよう修正。レグレッションテスト: `tests/spec/builtins.test.ts` Bug N2 ブロック。 | `a3fd929` |
| **Bug P: `AscB("")` が Error 5 にならず 0 を返す** | `AscB` に空文字列チェックを追加して Error 5 を投げるよう修正（`Asc`/`AscW` と同様）。 | `01c079a` |
| `eval("Exit Sub")` が JS 例外を漏らしてクラッシュ | `executeStatements` を try/catch でラップして Exit シグナルを飲み込む | `0ca97d8` |
| 同一ファイルへの二重 `Open` が Error 55 を出さない | `fileHandles` を走査して同一パスの重複チェックを追加 | `0ca97d8` |
| FSO `TextStream.AtEndOfStream` 未実装（Error 438） | `pos >= content.length` を返す getter を実装 | `0ca97d8` |

### 仕様準拠の動作（バグではない）

| 現象 | 説明 |
|---|---|
| `eval('m + 1')`（m は Long 変数）→ Error 424 | 実 VBA でも変数名 + 算術は文として `Call m(+1)` と解釈される。ワークアラウンド: `(m) + 1` |
| `eval('Dim x : x = 42 : x')` → `undefined` | マルチステートメント末尾の裸の識別子は Call 文扱い。値を得るには別の `eval('x')` で |
| `eval()` で他モジュールの `Dim`/`Private` 変数が読めない | `eval()` は独立したトップレベルモジュールとして評価されるため、他モジュールの `Private` 変数にはアクセスできない（VBA のクロスモジュール非公開変数と同じ意味論）。`Public` 変数はグローバル env 経由でアクセス可能。意図通りの設計。 |
| `On Error GoTo` ハンドラー内で `Exit Function` 後も `Err.Number` が残留 | `Function DivSafe(a,b): On Error GoTo H: DivSafe=a/b: Exit Function: H: DivSafe=0: End Function` → 呼出後 `Err.Number=11` | 実 VBA の仕様通り。VBA の `Err` はグローバルオブジェクトであり、`Resume` / `Resume Next` / `Err.Clear` / `On Error` 文の実行でのみリセットされる。`Exit Function` は `Err` をクリアしない。ワークアラウンド: エラーハンドラー内で明示的に `Err.Clear` を呼ぶ。 |

### 拡張機能 LSP のバグ（評価 #9 で発見・未修正）

| 問題 | 再現コード | 根本原因 |
|---|---|---|
| 引数付きチェーン補完が効かない | `ws.Cells(1, 1).` で Range でなくグローバル関数 29 件が返る | `detectMemberAccess` の正規表現が `)` 終端の式にマッチしない（`completion-provider.ts:454`） |
| VBA016 波下線が変数名を指す | `Dim x As UnknownType` で `x` に波下線（`UnknownType` であるべき） | `unknown-type-checker.ts:79` で `d.name.loc`（変数名位置）を渡している。パーサーの `VariableDeclarator` に型名の loc フィールドなし |

### ~~未修正バグ（評価 #21 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 21-1: `Dim s As String * N` がパースエラー**~~ | **修正済み**: `VariableDeclarator` に `fixedLength?: number` を追加。`parseDimStatement` が `As String * <整数>` の `* N` を消費して記録。`Environment.coerceToType` の `'String'` ブランチで `fixedLength` がある場合にパディング・切り捨てを適用。初期値は `'\0'.repeat(N)`（VBA 仕様準拠）。UDT `TypeMember` にも `fixedLength` を追加し `instantiateType` と MemberExpression 代入で同様の処理を実装。レグレッションテスト: `tests/spec/fixed-length-string.test.ts`（12 テスト）。 | |

### ~~未修正バグ（評価 #22 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 22-1: Variant 変数に整数を代入すると TypeName/VarType が subtype を失い Double を返す**~~ | **修正済み（評価 #22）**: `evaluator.ts:evaluateTypeIntrinsic` の非リテラル数値分岐に `Number.isInteger(val)` + Integer 範囲（-32768..32767）/Long 範囲チェックを追加。`Dim v As Variant : v = 42 : TypeName(v)` → "Integer" が正しく返るようになった。`VarType(v)` → 2（vbInteger）も正常。浮動小数点 `v = 3.14` → "Double" は変わらず正しい。レグレッションテスト: `tests/spec/typename.test.ts` Bug 22-1 ブロック。 | `evaluator.ts:5471`: TypeName/VarType の数値分岐が `argExpr.type === 'NumberLiteral'` のときだけ `inferLiteralTypeName` を呼び、それ以外（Variant 変数・引数など）は `return 'Double'` にフォールバックしていた。内部では JS `number`（64bit float）として格納されるため subtype タグが消える。 |

### 未修正バグ（評価 #12 で発見）

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~`__mocks__` 内の複数 VBA ファイルに同名関数があると "Ambiguous procedure" になる~~ | **修正済み**: `Environment.promoteProceduresFromModule` を追加し `loadVbaMock` で呼び出すことで後勝ち動作を実現。`A.bas` → `B.bas` の順でロードすると `B.bas` の定義が有効になる。 | |

### ~~未修正バグ（評価 #13 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~`parseAsClass` で生成した AST の `loc` が `undefined` → `.cls` ファイルで DefinitionProvider/ReferencesProvider/RenameProvider が機能しない~~ | **修正済み（評価 #13）**: `parseClassBody()` 内で各文パース後に `tok`（ブランチ開始前のトークン）と `this.tokens[this.pos-1]`（パース後の最終トークン）から `loc` を設定するよう修正。`ClassDeclaration` 本体の `loc` も最初・最後の文トークンから設定。`tsc -b` / `class-module.test.ts` 全通過。 | `parser.ts:parseClassBody()` が `parseProcedureDeclaration()` を直接呼ぶため `parseStatement()` の `stmt.loc = { start, end }` 設定（line 1358）が実行されない。`ClassDeclaration` 本体も `{ type, name, fields, procedures, body }` のみで `loc` なし。`buildScopedSymbolTable` 内の `if (!proc.loc) continue;` ガードで全シンボルがスキップされる。影響: DefinitionProvider は常に `null`、ReferencesProvider はスコープ絞り込みと `includeDeclaration` 除外が無効、RenameProvider は全ファイル無差別テキスト置換。`SymbolProvider` は別の fallback を持ち動作するが位置情報がすべて `(0,0)` になる。 |

### 未修正バグ（評価 #14 で発見）

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~`CodeLensProvider`: `Test_*` プロシージャが常に `✓ Tested` を返す（疑似陽性）~~ | **修正済み**: `testProcReferences` の `start` を `proc.loc.start.line - 1`（宣言行含む）から `proc.loc.start.line`（宣言行スキップ）に変更。 | |

### ~~未修正バグ（評価 #15 で発見・評価 #16 で修正確認）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~`autoParensEdit` が戻り型付き Function を検出しない~~ | **修正済み（評価 #16 で確認）**: `PROC_NO_PARENS` 正規表現に `(?:\s+As\s+\w+)?` が追加されており `autoParensEdit('Function GetValue As Long')` → `{ insertCol: 17 }` が正常に返る。 | |
| ~~`TestRunner.runTests()` がスタブ実装（常に `passed` を返す）~~ | **修正済み（評価 #16 で確認）**: `runTests` の引数は `stmts` から `src: string` に変更され、内部で `Evaluator` が実際に評価を行う本実装になっている。`Err.Raise` → `{ state: 'failed', message: 'Intentional failure' }` が正しく返る。 | |
| ~~`runTestWithEvaluation` のエラーメッセージが `"[object Object]"`~~ | **修正済み（評価 #16 で確認）**: `catch` 節が `(testError as any)?.message ?? String(testError)` になっており、VBA エラー plain object の `.message` プロパティが正しく取り出される。`Err.Raise 1, , "Intentional failure"` → `result.message === "Intentional failure"` が確認できた。 | |

### ~~未修正バグ（評価 #17 で発見・修正済み）~~

| 問題 | 最小再現コード | 修正コミット |
|---|---|---|
| ~~`Const` ホバーで型と値が表示されない~~ | **修正済み（`03f2b23`）**: `parser.ts` で `objectType` を `ConstDeclaration` に保持し、`symbol-table.ts` の `constLiteralText()` ヘルパーで値を文字列化。`Const MAX As Long = 100` と表示されるようになった。 | `03f2b23` |
| ~~`Public`/`Private` キーワードが hover で lowercase 表示~~ | **修正済み（`03f2b23`）**: `symbol-table.ts` に `cap()` ヘルパーを追加し、変数・定数・プロシージャの scope 表示を capitalize するよう修正。`Public count As Long` と正しく表示されるようになった。 | `03f2b23` |
| ~~`buildExtractFunctionEdit` が Dim 宣言とパラメーターを共存させる不正コードを生成する（Bug R1）~~ | **修正済み（評価 #18）**: 選択範囲に `Dim x` が含まれ `x` が inputs/outputs パラメーターになる場合、`reindented` ステップで `Dim x` が残存し `ByRef x As Variant` と重複する VBA コンパイルエラーが発生していた。`paramVarNames` set を使って `Dim` 行をフィルタリングする修正を `server.ts:819-826` に適用。`lsp-code-actions.test.ts` Test 13 でレグレッションテスト追加。 | 評価 #18 |
| ~~With ブロック内でユーザー定義クラスの Property Get が Error 424（Bug W1）~~ | **修正済み（評価 #19）**: `evaluateImplicitWithObjectExpression`（`evaluator.ts`）に `__vbaClass__` ブランチが欠如。VBA クラスを With オブジェクトとした場合、`.PropertyGet` が JS オブジェクトのメンバーとして探索されて失敗。`evaluateMemberExpression` と同じ getter → callClassMethod パターンを追加。`.Summary()`（Function 呼び出し）は別コードパスで動作していたため非対称になっていた。レグレッションテスト: `tests/spec/class-module.test.ts` 末尾 Bug W1 ブロック。 | 評価 #19 |
| ~~UDT 変数（`Dim pt As Point`）の `pt.` 補完が 0件（Bug LSP-1）~~ | **修正済み（評価 #20）**: `completion-provider.ts` の `getMembersForType()` が `ClassDeclaration` のみ探索し `TypeDeclaration` を未対応だった。`TypeDeclaration` ブランチを追加し、メンバーを `CompletionItemKind.Field` で返すよう修正。レグレッションテスト: `lsp-completion.test.ts` Test 14。 | 評価 #20 |
| ~~Enum 定数のホバーが null（Bug LSP-2）・Type 名のホバーが null（Bug LSP-3）~~ | **修正済み（評価 #20）**: `symbol-table.ts` の `collectScopedSymbols()` が `TypeDeclaration` / `EnumDeclaration` を未対応だった。`SymbolKind` に `'udt'` / `'enum-member'` を追加し、`TypeDeclaration` は型名を `Type Point (X As Long, Y As Long)` 形式で、`EnumDeclaration` は各メンバーを `Direction.North = 1` 形式でシンボルテーブルに登録。`hover-provider.ts` の `kindContextLabel()` にも対応 case を追加。レグレッションテスト: `lsp-hover.test.ts` 末尾 Bug LSP-2/3 ブロック。 | 評価 #20 |

### ~~未修正バグ（評価 #27 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 27-1: `CallByName` の `VbLet(4)` / `VbSet(8)` 未実装**~~ | **修正済み**: `builtins.ts` の `callbyname` に VbLet(4)/VbSet(8) ブランチを追加。`classDef.procedures` から `propertyType === 'let'`/`'set'` の手続きを検索して `ctx.callMethod` で呼び出す。fallback として逆タイプ（Let→Set/Set→Let）も試みる。レグレッションテスト: `tests/spec/callbyname.test.ts`（4テスト）。 | `builtins.ts:155-156` の `callbyname` が `VbGet(2)`/`VbMethod(1)` しかハンドルせず、それ以外は即 `throwError`。VBA クラスの Property Let を検索・呼び出すブランチが未実装。 |

### ~~未修正バグ（仕様調査 #27-続 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug A: `Replace(s, find, repl, start, count, compare)` — start/count/compare 引数が Error 450**~~ | **修正済み**: `builtins.ts` の `replace` を完全再実装。`start` は結果が prefix を除いたスライス、`count=-1` で無制限、`compare=1`（vbTextCompare）で大文字小文字無視。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug A ブロック（8テスト）。 | 旧実装は `(s: any, f: any, r: any)` の3引数固定登録。`start`/`count`/`compare` のパラメーター登録がなく Error 450 になっていた。 |
| ~~**Bug B: `Weekday(date, firstdayofweek)` — firstdayofweek 引数が Error 450**~~ | **修正済み**: `builtins.ts` の `weekday` 登録に `FirstDayOfWeek`（省略可）を追加。`fdow=0` は仕様通り `1`（vbSunday）扱い、`weekStart = fdow <= 1 ? 0 : fdow - 1` で JS 0-based 曜日にマッピング。レグレッションテスト: `tests/spec/datetime.test.ts` Bug B ブロック（5テスト）。 | 旧実装は `(d: any)` 単引数固定登録。 |
| ~~**Bug C: `DateDiff(interval, d1, d2, firstdayofweek, firstweekofyear)` — 第4/5引数が Error 450**~~ | **修正済み**: `builtins.ts` の `datediff` 登録に `FirstDayOfWeek`/`FirstWeekOfYear`（省略可）を追加。`"ww"` インターバルは両日付を週境界に丸めてから差を計算（`weekStart` オフセット反映）。レグレッションテスト: `tests/spec/datetime.test.ts` Bug C ブロック（4テスト）。 | 旧実装は `(interval, date1, date2)` の3引数固定登録。 |
| ~~**Bug D: `Format(expr, fmt, firstdayofweek, firstweekofyear)` — 第3/4引数が Error 450**~~ | **修正済み**: `builtins.ts` の `format` 登録パラメーター配列に `FirstDayOfWeek`/`FirstWeekOfYear`（省略可）を追加（実装本体は変更なし・引数は受けるが無視）。`Format(Now(), "dddd", 2)` が Error 450 なく動作するようになった。 | `format` の `ctx.reg` 呼び出しのパラメーター配列が2要素固定（`Expression`/`Format`）だった。 |
| ~~**Bug E: `RGB(r, g, b)` が Error 35（未実装）**~~ | **修正済み**: `builtins.ts:registerConstants` に `rgb` を登録。COLORREF 形式 `r + g*256 + b*65536`。各チャネルは 0-255 にクランプ。レグレッションテスト: `tests/spec/builtins.test.ts` RGB/QBColor/Nz ブロック。 | 定数テーブルは vbRed 等のエイリアスを持たず、`RGB` 関数自体が `builtins.ts` に未登録だった。 |
| ~~**Bug F: `QBColor(n)` が Error 35（未実装）**~~ | **修正済み**: `builtins.ts:registerConstants` に `qbcolor` を登録。16色テーブルで 0〜15 を COLORREF に変換。範囲外は Error 5。レグレッションテスト: `tests/spec/builtins.test.ts` RGB/QBColor/Nz ブロック。 | `builtins.ts` に `qbcolor` の登録が存在しなかった。 |
| ~~**Bug G: `Nz(value, valueifnull)` が Error 35（未実装）**~~ | **修正済み**: `builtins.ts:registerConstants` に `nz` を登録。`Null`/`Empty`/`null`/`undefined` なら `valueifnull`（省略時は `0`）を返す。レグレッションテスト: `tests/spec/builtins.test.ts` RGB/QBColor/Nz ブロック。 | Access VBA の組み込み関数で Excel VBA には存在しないため未登録だった。 |
| ~~**Bug H: `StrConv(s, conv, LCID)` 第3引数が Error 450**~~ | **修正済み**: `strconv` 登録に `{ name: 'LCID', optional: true }` を追加（LCID は無視するが引数を受け取るようになった）。レグレッションテスト: `tests/spec/builtins.test.ts`。 | `ctx.reg` の params 配列が `String`/`Conversion` の2要素固定だった。 |
| ~~**Bug I: `InputBox(prompt, title, default, XPos, YPos, ...)` 第4引数以降が Error 450**~~ | **修正済み**: `inputbox` 登録に `XPos`/`YPos`/`HelpFile`/`Context`（すべて省略可）を追加（スタブ実装・無視）。 | params 配列が3要素固定だった。 |
| ~~**Bug J: `MsgBox(prompt, buttons, title, HelpFile, Context)` 第4引数以降が Error 450**~~ | **修正済み**: `msgbox` 登録に `HelpFile`/`Context`（省略可）を追加（スタブ実装・無視）。 | params 配列が3要素固定だった。 |
| ~~**Bug K: `CreateObject(class, ServerName)` 第2引数が Error 450**~~ | **修正済み**: `createobject` 登録に `{ name: 'ServerName', optional: true }` を追加（無視）。 | params 配列が1要素固定だった。 |
| ~~**Bug L: `Format(True, "Yes/No")` が "Yes" でなく "True" を返す**~~ | **修正済み**: `formatFunc` の named format ブランチで `VbaBoolean` を検出して `yes/no`→"Yes"/"No"、`on/off`→"On"/"Off"、`true/false`→"True"/"False" を返すよう修正。VbaBoolean 以外の数値型は `val.value` を `formatNumber` に渡す。レグレッションテスト: `tests/spec/builtins.test.ts`。 | `VbaBoolean` は `typeof 'object'` のため `typeof val === 'number'` が false になり `String(val)` → "True"/"False" へフォールバックしていた。 |
| ~~**Bug M: `vbFirstJan1`/`vbFirstFourDays`/`vbFirstFullWeek`/`vbDecimal`/`vbDataObject`/`vbUserDefinedType` が未登録（`Null` を返す）**~~ | **修正済み**: `registerConstants` に 6 定数を追加（`vbFirstJan1=1`/`vbFirstFourDays=2`/`vbFirstFullWeek=3`/`vbDecimal=14`/`vbDataObject=13`/`vbUserDefinedType=36`）。未登録定数を参照すると暗黙的に `Empty` が返り、`firstweekofyear` 引数として渡すと 0 扱いで誤動作する。 | VarType 定数ラインに `vbDataObject`/`vbDecimal`/`vbUserDefinedType` が含まれておらず、firstweekofyear 定数ライン自体が存在しなかった。 |
| ~~**Bug N: `Left(Null, n)` / `Right(Null, n)` / `Mid(Null, n)` が Null でなく文字列を返す**~~ | **修正済み**: 各関数の先頭に `if (val === vbaNull) return vbaNull;` を追加。`Left(Null, 2)` → `Null`（`IsNull=True`）が正しく返るようになった。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug N ブロック。 | `leftFunc`/`rightFunc`/`midFunc` で `String(val ?? '')` を先に評価しており、`vbaNull` は Symbol のため `??` で素通りして `String(Symbol(vbaNull))` → `"Symbol(vbaNull)"` になっていた。 |

### ~~未修正バグ（評価 #28 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 28-1: `ReDim Preserve` で UDT 配列を拡張後、新インデックス要素のメンバーアクセスが Error 424**~~ | **修正済み**: `evaluateReDimDeclarator` の Preserve パス内に `fillArrayWithUDT(arr, dims, 0, typeName, true)` 呼び出しを追加。`skipExisting=true` のとき既存要素をスキップし、`undefined`/`0` のスロットのみ `instantiateType` で初期化する。レグレッションテスト: `tests/spec/array-functions.test.ts` Bug 28-1 ブロック。 | `evaluator.ts:5172-5174` の `if (!isPreserve)` ガードにより `fillArrayWithUDT` が Preserve 時にスキップされ、`copyPreservedData` は旧インデックスのみコピーするため新インデックスの UDT 要素が `0` のまま残っていた。Long/String の通常配列では `0`/`""` を直接メンバーアクセスせず代入で済むため問題が顕在化しなかった。 |

### ~~未修正バグ（評価 #29 で発見・全件メイン再現確認済み）~~ → 全件修正済み

| 問題 | 修正ファイル | テスト |
|---|---|---|
| **Bug 29-A: 型サフィックス付き宣言 `Dim n&` が機能しない** | `parser.ts: parseDimStatement / parsePrimary` | `tests/spec/type-system.test.ts` |
| **Bug 29-B: `^` が右結合（VBA は左結合）** | `parser.ts: parseExponentiation` | `tests/spec/operators-extra.test.ts` |
| **Bug 29-C: 負の底の非整数べき乗が Error 5 にならない** | `evaluator.ts: case '^'` | `tests/spec/operators-extra.test.ts` |
| **Bug 29-D: `Null & Null` が Null にならない** | `evaluator.ts: op === '&'` | `tests/spec/operators-extra.test.ts` |
| **Bug 29-E: `Array()` が `Option Base 1` を無視する** | `builtins.ts: array 登録 / UBound` | `tests/spec/option_base.test.ts` |
| **Bug 29-F: `eval()` の裸の引数なしメソッド呼び出しが静かに no-op** | `evaluator.ts: evaluateMemberExpression` | `tests/spec/class-module.test.ts` |
| **Bug 29-G: eval 複文中の裸引数なしメソッドが Error 450** | `parser.ts: call arg check (OperatorColon 除外)` | `tests/spec/class-module.test.ts` |
| **Bug 29-H: BEGIN/END なしの部分 .cls ヘッダーが全メンバー Error 438** | `preprocessor.ts: stripVBAFileHeader` | `tests/spec/preprocessor-cls-header.test.ts` |

### ~~未修正バグ（評価 #33 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 33-A: クラス内の非修飾・括弧なし自メンバー参照（Property Get / Function）が silent Empty**~~ | クラス内 `CallsProp = PropX & "!"` → `"!"`（Option Explicit でも検出されない） | クラスメンバーは env のプロシージャマップに載らず、未定義変数の暗黙 Empty に解決されていた。Identifier 評価に「変数にもプロシージャにも解決できない場合、Me のクラスメンバー（Function/Property Get、必須引数 0）を暗黙の Me.<name> として呼ぶ」フォールバックを追加。`env.get()` の暗黙初期化より先に `hasVariable` を取るのが要点。`class-module.test.ts` Bug 33-A ブロック |
| ~~**Bug 33-B: `Collection.Add` の名前付き `Before:=`/`After:=` が誤バインド**~~ | `c.Add "C", After:=1` → 末尾追加（値が Key 位置に化ける） | `VbaCollection.add` に `__vbaParamSpec__` がなく、名前付き引数の値が位置引数に落ちていた。仕様を付与し、さらに仕様にない名前付き引数・仕様なし関数への名前付き引数は Error 448 で明示的に失敗させるように。`collection.test.ts` Bug 33-B ブロック |
| ~~**Bug 33-C: 未使用の `Dim x As New Class` を ByVal で渡すとメンバー変更が呼び出し元に見えない**~~ | `Dim b As New Widget : Poke b : b.Cnt` → 0（期待 1） | auto-instance プレースホルダーが callee 側で実体化され、caller の変数に書き戻されなかった。ユーザープロシージャ呼び出しの引数評価（位置・名前付きとも）で `resolveAutoInstance` を通し、渡す時点で caller 側に実体化するよう修正。`auto-instance-args.test.ts` Bug 33-C ブロック |

### ~~未修正バグ（評価 #32 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 32-A: `Tab(n)` が n+1 桁目に出力（オフバイワン）**~~ | `Debug.Print "X"; Tab(10); "Y"` → Y が 11 桁目 | Print #/Debug.Print の Tab 処理が `n - output.length` でパディングしていた（正しくは `n - 1 - output.length`）。両経路とも修正。`tests/spec/write-input-print-zones.test.ts` |
| ~~**Bug 32-B: `Write #1, #2024/03/15#` がパースエラー**~~ | 同左 → `Parse error: unexpected token '2024'` | レクサーの日付リテラル判定がカンマ許容の緩い正規表現で `#1, #` を日付と誤認。日付・時刻区切り（`/` `-` `:`）の存在を必須化して修正。`lexer.ts` |
| ~~**Bug 32-C: `Input #` が引用符内カンマで分割**~~ | `Write #f, "comma, inside", 42` → `Input #f, s, n` が Error 13 | 単純 `split(",")` だった。引用符状態を追跡するフィールド分割に変更。 |
| ~~**Bug 32-D: `Write #` の日付書式と改行が実 VBA と不一致**~~ | 出力が `#2024/03/15#` + LF（期待: `#2024-03-15#` universal format + CRLF） | 日付を universal format で書き、改行を CRLF に統一（Print # と対称に）。`Input #` にも `#日付#` フィールドの VbaDate 復元を追加。 |
| ~~**Bug 32-E: `LSet udtB = udtA`（UDT 間コピー）が Error 424**~~ | 同レイアウト Type 2 つで `LSet b = a` | LSet が文字列パスしかなかった。同一レイアウト（フィールド数一致）なら位置ベースでフィールドコピー、不一致なら明示的な Error 5 を実装。`tests/spec/lset-rset.test.ts` Bug 32-E ブロック |

### ~~未修正バグ（評価 #31 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 31-A: `Public/Private Static Sub|Function` がパースエラー**~~ | `Public Static Function F() As Long` → `Parse error: Expected variable name (Found Static)` | `parser.ts` の `KeywordPublic/Private` 分岐が `Static` をチェックせず `parseDimStatement` にフォールスルーしていた。`Static` + `Sub|Function|Property` の先読み分岐を追加して修正。レグレッションテスト: `tests/spec/static.test.ts` Bug 31-A ブロック。 |
| ~~**Bug 31-B: `Erl` が未実装で常に 0 を返す（黙って識別子暗黙解決）**~~ | `On Error GoTo H` + `20 Err.Raise 5` + ハンドラー内 `Erl` → 0（期待 20） | `erl` の登録がどこにもなく、未定義識別子として暗黙 0 に解決されていた。数値行ラベル通過時に `lastLineNumberLabel` へ記録し、エラー捕捉時に `Err` 状態へ確定する実装を追加。`Err.Clear`/`Resume`/`On Error` 文でリセット。レグレッションテスト: `tests/spec/erl-function.test.ts` Bug 31-B ブロック（既存の INFO-only テスト 1・2 もハードパスに）。 |

### ~~未修正バグ（評価 #30 で発見・修正済み）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 30-A: `VarType(classInstance)` が 36 (vbUserDefinedType) を返す（仕様: 9 = vbObject）**~~ | `Dim obj As New MyClass : VarType(obj)` → 36 | `builtins.ts` の `vartype` 登録で `if (val.__vbaTypeName__) return 36` が `if (val.__vbaClass__) return 9` より先に実行されていた。クラスインスタンスは両方のプロパティを持つため UDT 扱いになっていた。チェック順を入れ替えて修正。`TypeName(obj)` は正常動作していた。レグレッションテスト: `tests/spec/typename.test.ts` Bug 30-A ブロック。 | `builtins.ts:109-110` の順序を入れ替え |

### 未修正バグ（評価 #26 で発見）

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 26-1: `Put #n,,var` はバイナリでなくテキスト書き込み**~~ | **評価 #36 で修正済み**: `Byte`/`Integer`/`Long` と CP932 `String` を型サイズ・バイト列で書き込む。 | |
| ~~**Bug 26-2: `Get #n,,var` は常に1024バイト一括読み込み（型サイズ無視）**~~ | **評価 #36 で修正済み**: 宣言型・固定長文字列長から必要な範囲だけを読み、ハンドル位置を消費バイト数分だけ進める。 | |
| ~~**Bug 26-3: `Open path For Random As #n Len = recLen` がパースエラー**~~ | **修正済み**: `parser.ts:parseOpenStatement` でファイル番号消費後に `peek().type === Identifier && value.toLowerCase() === 'len'` + `OperatorEquals` を消費して `recordLen` を記録するよう修正。レグレッションテスト: `tests/spec/filesystem.test.ts` Bug 26-3 ブロック。 | |
| ~~**Bug 26-4: UDT変数を `Put` すると "[object Object]" を書き込む**~~ | **評価 #36 で修正済み**: 配列を含まない UDT を宣言順に連続したバイナリフィールドとして書き込む。 | |
| ~~**Bug 26-5: UDT変数への `Get` で Error 424**~~ | **評価 #36 で修正済み**: 同じ範囲の UDT を復元して代入する。 | |
| ~~**Bug 26-7: `GetAttr(path)` が Error 35（未実装）**~~ | **修正済み**: `evaluator.ts` に `getattr`（常に 0 を返すスタブ）と `setattr`（no-op）を登録。`builtins.ts:registerConstants` に `vbNormal`/`vbReadOnly`/`vbHidden`/`vbSystem`/`vbVolume`/`vbDirectory`/`vbArchive`/`vbAlias` 定数を追加。レグレッションテスト: `tests/spec/filesystem.test.ts` Bug 26-7 ブロック。 | |

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
| ~~`Dim s As String * N`（固定長文字列）が未実装~~（Bug 21-1）~~ | **修正済み**: `VariableDeclarator.fixedLength` と `TypeMember.fixedLength` を追加。`parseDimStatement` / Type ブロックパーサーで `As String * N` の `* N` を消費・記録。`Environment.coerceToType` でパディング・切り捨てを適用。UDT メンバー代入も `__fixedLengths__` で対応。`tests/spec/fixed-length-string.test.ts`（12 テスト）。 |
| ~~`Currency` 型が固定小数点演算でない~~ | **修正済み（評価 #23 で確認）**: `CCur(0.1) + CCur(0.2) = 0.3` 厳密一致。内部 BigInt ベースの 4桁固定小数点演算に刷新済み。`VBARunner.run()` / `.eval()` の戻り値も `number` に正規化（Bug C-1 修正）。 |
| `Dim x As`（型名欠落）が構文エラーにならない | `Dim x As\n` と書くと parser の error recovery が黙って回復し、`VBA014` 未使用変数警告のみが出る。構文ミスを新規ユーザーが気付けない可能性がある。 |
| VBA003（ByRef/ByVal 省略）警告が severity:Warning で新規ユーザーに noisy | `Function Add(a As Long, b As Long)` のような標準的な宣言でも `VBA003` が severity 2（Warning）で出る。VBA の慣習では省略が普通のため、初回ロード時に「いきなり Warning が多い」印象を与えやすい。Hint（severity 4）への変更か設定で off 可能にすると親切。 |
| ~~`vba-types.json` 削除時に型スタブがリセットされない~~ | **修正済み**: `server.ts` に `clearTypeStubs()` を追加し、`extension.ts` の `typeStubsWatcher.onDidDelete` でフック。`vba-types.json` 削除時に補完プロバイダーの型スタブが即座にクリアされる。 |
| ~~LSP: 引数付きチェーン補完 `obj.Method(args).` が効かない~~ | **修正済み（評価 #10 で確認）**: `ws.Cells(1, 1).` で 48 件の Range メンバーが正しく返るようになった。`detectMemberAccess` の正規表現が `)` 終端を処理できるよう拡張済み（`completion-provider.ts:454`）。 |
| ~~LSP: シグネチャヘルプが文字列リテラル入力中に消える~~ | **修正済み**: `findCallContext` の文字列境界判定を Lexer トークンベースに変更。右→左スキャンの誤判定を解消。`findCallContext('Format(x, "', 11)` → `{ name: 'Format', activeParameter: 1 }` が正しく返る。 |
| ~~LSP: VBA016 診断の range が型名でなく変数名を指す~~ | **修正済み（評価 #10 で確認）**: `Dim x As UnknownType` で `column: 9, endColumn: 20`（UnknownType の位置）が返るようになった。`unknown-type-checker.ts:79` で `d.objectTypeLoc ?? d.name.loc` の優先順位が機能している。パーサーが `objectTypeLoc` を AST に格納するよう修正済み。 |
| `Dim empty As String` がパースエラー | `empty` は VBA 仕様上の予約語のため変数名に使えない。VBA 仕様準拠の正しい動作。 |
| ~~LSP: With ブロック内で引数付きメソッドチェーン後の補完が 0 件~~ | **修正済み（評価 #11 で確認）**: `    .Cells(1, 1).` でトリガーして 48 件の Range メンバーが返るようになった。コミット `3300dcb`。 |
| ~~時刻のみの日付リテラル未対応~~ | **修正済み** (`b4d00c3`): `#12:30:45#` / `#8:30:00 AM#` が Error 13 になっていた。`parseDateLiteral` が時刻のみの場合に基準日（1899/12/30）を返すよう修正。 |
| ~~`Class_Terminate` が参照カウントなしで早期発動~~ | **修正済み**: `Set p1 = Nothing` を呼んでも別変数 `p2` が同じオブジェクトを保持していれば `Class_Terminate` を呼ばないよう、参照カウント（`__refCount__`）を実装。`Set` 代入で addRef、`Set = Nothing` で releaseRef、スコープ脱出でもカウントを減算。`circular-reference-terminate.test.ts` 全 16 テスト通過。 |
| ~~**Bug R: ISO 日付文字列 "YYYY-MM-DD" のタイムゾーンオフセット混入**~~ | **修正済み**: `vba-types.ts:parseVbaDate` で ISO 形式（`/^\d{4}-\d{2}-\d{2}$/`）を検出し、`-` を `/` に置換してから `new Date()` に渡すよう修正。JS の ISO 形式は UTC midnight として解析されるため、UTC+9 環境では `getHours()=9` → VBA シリアル値に 0.375 が混入していた。スラッシュ形式は LOCAL midnight として解析されるため修正される。レグレッションテスト: `tests/spec/cdate.test.ts` Bug R ブロック。 | `f4c3856` |
| ~~**Bug S: `LenB(Null)` が 30 を返す（vbaNull Symbol を文字列化した長さ）**~~ | **修正済み**: `builtins.ts:lenb` に `if (s === vbaNull) return vbaNull;` を追加。`Len(Null)` は既に修正済みだったが `LenB` が未対応だった。レグレッションテスト: `tests/spec/builtin-strings.test.ts` Bug S ブロック。 | `c6544f7` |
| ~~**Bug T/U: `LeftB/RightB/MidB(Null)` が Null でなく文字列を返す**~~ | **修正済み**: 各関数の先頭に `if (val === vbaNull) return vbaNull;` を追加。`Left/Right/Mid` と同様の Null 伝播を実装。`tests/spec/builtin-strings.test.ts` Bug S ブロックで検証。 | `c6544f7` |
| ~~**Bug V/W: `Asc/AscW/AscB(Null)` が 83 を返す（vbaNull Symbol の先頭バイト）**~~ | **修正済み**: `ascFunc`/`ascb` の先頭に `if (s === vbaNull) return vbaNull;` を追加。`String(vbaNull ?? '')` が `"Symbol(vbaNull)"` を生成して `charCodeAt(0)=83` を返していた。 | `4f2172a` |
| ~~**Bug X: `Replace(Null, find, repl)` が "Symbol(vbaNull)" を返す**~~ | **修正済み**: `replace` の先頭に `if (s === vbaNull) return vbaNull;` を追加。VBA 仕様通り第1引数 Null → Null 伝播。 | `4f2172a` |
| ~~**Bug AA: `Space(Null)` が TypeError（Symbol→数値変換エラー）**~~ | **修正済み**: `spaceFunc` の先頭に `if (n === vbaNull) return vbaNull;` を追加。`Number(vbaNull)` が JS TypeError を投げていた。 | `2221d61` |
| ~~**Bug AB: `String(n, Null)` / `String(Null, char)` が誤った文字列を返す**~~ | **修正済み**: `stringFunc` の先頭に `if (n === vbaNull \|\| char === vbaNull) return vbaNull;` を追加。 | `2221d61` |
| ~~**Bug Y: `Format(文字列, 日付パターン)` が文字列フォーマットとして処理される**~~ | **修正済み**: `formatFunc` で `typeof effectiveVal === 'string'` 分岐に `isDatePattern` チェックを移動。`isDatePattern && !/^[0#,.%]+$/.test(fmt)` のとき `parseVbaDate` で変換してから `formatDate` を呼ぶよう修正。`Format("2024/03/15", "yyyy")` → "2024" が正しく返る。レグレッションテスト: `tests/spec/builtins.test.ts` Bug Y ブロック。 | `02a682d` |
| ~~**Bug Z: `Format(time, "hh:nn AM/PM")` が 24 時間制のままになる**~~ | **修正済み**: `format.ts:formatDate` でトークン列を事前スキャンして `am/pm`/`ampm`/`a/p` マーカーがある場合は `use12Hour=true` にし、`hh`/`h` トークンを `h12`（12 時間制）で返すよう修正。`Format(CDate("14:30:00"), "hh:nn AM/PM")` → "02:30 PM" が正しく返る。レグレッションテスト: `tests/spec/builtins.test.ts` Bug Z ブロック。 | `f34b9b0` |
| ~~**Bug AC: `Replace(str, Null, repl)` / `Replace(str, find, Null)` が Null でなく文字列を返す**~~ | **修正済み**: `builtins.ts:replace` の先頭条件を `if (s === vbaNull \|\| f === vbaNull \|\| r === vbaNull) return vbaNull;` に拡張。find が Null だと `String(vbaNull)="Symbol(vbaNull)"` に変換されて検索対象が見つからず元文字列がそのまま返っていた。repl が Null だと "Symbol(vbaNull)bc" が返っていた。 | `51b32af`〜 |
| ~~**Bug AD: `Null Like pattern` / `str Like Null` が Null でなく True/False を返す**~~ | **修正済み**: `evaluator.ts` の `case 'like':` に `if (leftVal === vbaNull \|\| rightVal === vbaNull) return vbaNull;` を追加。VBA の Like 演算子はいずれかのオペランドが Null なら Null を返す。 | `51b32af`〜 |
| ~~**Bug AE: `Chr(Null)` / `ChrW(Null)` が TypeError でクラッシュする**~~ | **修正済み**: `chrFunc`/`chrwFunc` の先頭に `if (n === vbaNull) return vbaNull;` を追加。`Number(Symbol)` が JS TypeError を投げていた。 | `51b32af`〜 |
| ~~**Bug AF: `Hex(3.7)` / `Oct(3.7)` が切り捨てではなく四捨五入されない**~~ | **修正済み**: `hexFn`/`octFn` で `Math.floor` を `vbaRound(..., 0)` に変更。VBA は整数変換にバンカーズ丸め（偶数丸め）を使うため `Math.floor` では不正確。`Hex(3.7)` = "4", `Hex(2.5)` = "2", `Hex(3.5)` = "4" が正しく返る。 | `51b32af`〜 |
| ~~**Bug AG/AH/AI: `DateAdd/DateDiff/DatePart` の date 引数が Null のとき TypeError でクラッシュする**~~ | **修正済み**: 各関数の先頭に `if (date === vbaNull) return vbaNull;`（DateDiff は date1/date2 両方）を追加。`parseVbaDate(Symbol)` → `String(Symbol)` → `Number(Symbol)` で JS TypeError が発生していた。 | `51b32af`〜 |
| ~~**Bug AJ: `ChrB(Null)` が TypeError でクラッシュする**~~ | **修正済み**: `chrb` 関数の先頭に `if (n === vbaNull) return vbaNull;` を追加。`Number(Symbol)` で JS TypeError が発生していた。 | `51b32af`〜 |
| ~~**Bug AK: `Year/Month/Day/Hour/Minute/Second/Weekday(Null)` が TypeError でクラッシュする**~~ | **修正済み**: 各関数の先頭に `d === vbaNull ? vbaNull :` チェックを追加。`parseVbaDate(Symbol)` → `String(Symbol)` で JS TypeError が発生していた。 | `51b32af`〜 |
| ~~**Bug AL: `WeekdayName(Null)` / `MonthName(Null)` が TypeError でクラッシュする**~~ | **修正済み**: 各関数の先頭に `if (weekday/month === vbaNull) return vbaNull;` を追加。`Number(Symbol)` で JS TypeError が発生していた。 | `51b32af`〜 |
| ~~**Bug AM/AN: `Split(Null, ",")` が Null でなく `["Symbol(vbaNull)"]` を返す / `Join(arr, Null)` がクラッシュ**~~ | **修正済み**: `split` 先頭に `if (s === vbaNull) return vbaNull;`、`join` に `if (del === vbaNull) return vbaNull;` を追加。 | `51b32af`〜 |
| ~~**Bug AO: `Choose(Null, ...)` が Type mismatch エラー（VBA 仕様では Null を返す）**~~ | **修正済み**: `choose` の先頭の `ctx.throwError(TYPE_MISMATCH)` を `return vbaNull` に変更。VBA ドキュメント「index が Null の場合 Null を返す」に準拠。 | `51b32af`〜 |
| ~~**Bug AP/AQ/AR: `Left(str, Null)` / `Right(str, Null)` / `Mid(str, Null, n)` / `Mid(str, n, Null)` が TypeError でクラッシュ**~~ | **修正済み**: `leftFunc`/`rightFunc`/`midFunc` の Null チェックを `val === vbaNull \|\| len/start === vbaNull` に拡張。 | `51b32af`〜 |
| ~~**Bug AS/AT: `DateValue(Null)` / `TimeValue(Null)` が TypeError でクラッシュ**~~ | **修正済み**: `datevalue`/`timevalue` の先頭に `if (val === vbaNull) return vbaNull;` を追加。 | `51b32af`〜 |
| ~~**Bug AU/AV/AW: `LeftB(str, Null)` / `RightB(str, Null)` / `MidB` の Null 引数でクラッシュ**~~ | **修正済み**: `leftb`/`rightb`/`midbFunc` の Null チェックを `val \|\| len/start` に拡張。 | `51b32af`〜 |
| ~~**Bug AX/AY: `DateSerial(Null,...)` / `TimeSerial(Null,...)` が TypeError でクラッシュ**~~ | **修正済み**: 各関数の先頭に Null チェックを追加。`Number(Symbol)` で JS TypeError が発生していた。 | `51b32af`〜 |
| ~~**Bug AZ/BA/BB: `And`/`Or`/`Imp` が Null を含む場合に単純 Null 伝播を行い、三値論理の特例を無視**~~ | **修正済み**: `evaluateBinaryExpression` で Null チェック前に And/Or/Imp の特例を処理。`False And Null`=False、`True Or Null`=True、`False Imp Null`=True、`Null Imp True`=True。 | `51b32af`〜 |
| ~~**Bug BC: `FormatCurrency` / `FormatNumber` / `FormatPercent` / `FormatDateTime` が未実装**~~ | **修正済み**: 4関数を `builtins.ts` に追加。`FormatCurrency(1234.5)` = "$1,234.50"、`FormatPercent(0.5)` = "50.00%"、`FormatDateTime(date, namedFmt)` で vbGeneralDate(0)/vbLongDate(1)/vbShortDate(2)/vbLongTime(3)/vbShortTime(4) に対応。 | `077a26e`〜 |
| ~~**Bug BD: `CByte`/`CInt`/`CLng`/`CSng`/`CDbl`/`CDec`/`CCur`/`CLngLng`(Null) が Error 13 を返す（VBA では Error 94）**~~ | **修正済み**: 各変換関数の先頭に `if (val === vbaNull) ctx.throwError(INVALID_USE_OF_NULL, ...)` を追加。`CLngLng` は `TYPE_MISMATCH` から `INVALID_USE_OF_NULL` に変更。 | `c490d76`〜 |
| ~~**Bug BE: `Filter(arr, Null)` が空配列を返す（VBA では Error 13 Type mismatch）**~~ | **修正済み**: `filter` 関数に `if (match === vbaNull) ctx.throwError(TYPE_MISMATCH, ...)` を追加。`String(Symbol)` が `"Symbol(vbaNull)"` になる問題を回避。 | `c490d76`〜 |
| ~~**Bug BF: `Nz(Null)` が `0` を返す（VBA では `""` を返す）**~~ | **修正済み**: `nz` のデフォルト引数を `vbaMissing` に変更し、未指定時は `''` を返すよう修正。`Nz(Null, 42)` は引き続き `42` を返す。 | `c490d76`〜 |
| ~~**Bug BG: `Format(2.5, "0")` が `"3"` を返す（VBA では `"2"` — 銀行家丸め）**~~ | **修正済み**: `format.ts` の `formatSection` で `n.toFixed(maxDecimals)` を `vbaRound(n, maxDecimals).toFixed(maxDecimals)` に変更。`builtins.ts` の `fmtNumeric` (FormatCurrency/FormatNumber/FormatPercent) も同様に修正。`Format(2.5, "0")`="2"、`Format(3.5, "0")`="4"、`Format(1234.5, "0")`="1234"。 | `1a93ea9`〜 |
| ~~**Bug BH: `7.5 \\ 2` が `3` を返す（VBA では `4` — 被演算子を事前に整数化してから除算）**~~ | **修正済み**: `evaluator.ts` の `\\` と `Mod` 演算子に `_vbaRound(x, 0)` で各引数を整数化してから演算するよう修正。`7.5 \\ 2`=4、`7.5 Mod 2`=0。 | `f047068`〜 |
| ~~**Bug BI: `2 ^ -1` がパースエラーになる（VBA では 0.5）**~~ | **修正済み**: `parser.ts` の `parseExponentiation` で右辺の `parsePrimary()` を `parseUnary()` に変更し、単項 `-` が `^` の右辺に出現できるようにした。 | `f047068`〜 |
| ~~**Bug BJ: Currency混在の `\\`/`Mod` が事前丸めと 0 除算チェックを欠いていた**~~ | **修正済み**: `evaluateCurrencyOp` の整数パス（`bankersDivide` ベース）と浮動小数点フォールバックパスの両方で `_vbaRound(x / 10000, 0)` による事前丸めと `ri === 0` 判定を追加。`CCur(6.5) Mod 3`=0（旧: 5000）、`CCur(5.5) \\ CCur(3.5)`=1（旧: 2）が正しく返る。 | `04838c3`〜 |
| ~~**Bug BK: `CDate(Null)`/`Str(Null)`/`Val(Null)` が Error 94 の代わりに異常動作**~~ | **修正済み**: `builtins.ts` の各変換関数で `vbaNull` チェックを追加して `INVALID_USE_OF_NULL` を投げるよう修正。`CStr(Null)` は例外ではなく `""` を返すよう修正。 | `95e1840`〜 |
| ~~**Bug BL: `Input#` が `#TRUE#`/`#FALSE#`/`#NULL#` トークンを解析できなかった**~~ | **修正済み**: `evaluateInputStatement` に `parseInputValue` ヘルパーを追加し、`#TRUE#`→`vbaTrue`、`#FALSE#`→`vbaFalse`、`#NULL#`→`vbaNull`に変換。 | `a4b229c`〜 |
| ~~**Bug BM-1: `Write#` で `Empty` が `#NULL#` と書き出されていた**~~ | **修正済み**: `vbaEmpty===null`（JavaScript null）のため、`val===null` チェックが `vbaNull`（Symbol）より先に `vbaEmpty` も捕捉していた。チェック順を入れ替え `vbaEmpty` を先にチェックし `""`、`vbaNull` を後にチェックし `#NULL#` を返すよう修正。 | `2748e2a` |
| ~~**Bug BM-2: `Line Input#` がモジュールレベル変数に書き込めなかった**~~ | **修正済み**: `evaluateLineInputStatement` が `env.setLocally()` を使っていたため、プロシージャ終了後に変数が破棄されていた。`evaluateAssignmentToVariable()` に変更しスコープチェーンを正しく辿るよう修正。 | `2748e2a` |
| ~~**Bug BN: `Erase` 後の動的配列への `UBound`/`LBound` が Error 9 を投げなかった**~~ | **修正済み**: `evaluateEraseStatement` の動的配列パスで `d=[]` と空配列を代入していたが、`UBound([])=-1` で誤動作。`d=null` に変更し既存の `!Array.isArray` チェックで Error 9 が投げられるよう修正。`ReDim` は引き続き正常動作。 | `2748e2a` |
| ~~**Bug BO: `evalExpression("True + True")` が Error 424 になっていた**~~ | **修正済み**: `isCallableLeftmostLeaf` が VBA キーワード定数（`True`/`False`/`Null`/`Empty`/`Nothing`）の `Identifier` も callable とみなし statement fallback に飛ばしていた。定数名を除外する分岐を追加。`evalExpression('True + True')=-2`, `'False + 1'=1`, `'Null + 1'=Null` が正しく動作するようになった。 | `c987947` |
| ~~**Bug BP: `evalExpression("\"abc\" = \"ABC\"")` が `undefined` を返していた**~~ | **修正済み**: `isStatementAmbiguous` の `=` 演算子チェックが左辺の型に関係なくトリガーされ、文字列/数値リテラル同士の等値比較が statement fallback に飛んで結果が捨てられていた。修正: `=` も `+`/`-` と同様に `isCallableLeftmostLeaf` チェックを追加し、左辺が代入可能なノードの場合のみ ambiguous とみなすよう変更。代入セマンティクス（`x = 42`）は変わらず動作。 | `06a3ded` |
| ~~**Bug BQ: `Format(time, "Long Time")` が 24 時間表記で返っていた / `Format(date, "Short Date")` が `yyyy/MM/dd` 形式だった**~~ | **修正済み**: `format.ts` の named format テーブルで `"long time"` が `HH:mm:ss`（24h）、`"short date"` が `yyyy/MM/dd` になっていた。`"long time"` → `h:mm:ss AM/PM`（12h）、`"short date"` → `M/D/YYYY`（米国形式）に修正。 | *(このセッション)* |
| ~~**Bug BR: `Debug.Print "a"; "b"` が `ParseError: unexpected token ';'` で失敗していた**~~ | **修正済み**: `Debug.Print` は `CallStatement` として解析されていたため `;` セパレーターが不正なトークンとなっていた。パーサーに `DebugPrintStatement` 型を追加し、`Debug.Print` を `parsePrintStatement` と同等の print-list 構文（`;` / `,` / `Spc()` / `Tab()` 対応）で解析するよう変更。評価器に `evaluateDebugPrintStatement` を追加し `onPrint` コールバックに出力。 | `f038feb` |
| ~~**Bug CB: `Dim c As Color` (Enum 型変数) の `TypeName(c)` が "Double" を返していた**~~ | **修正済み**: `evaluateVariableDeclaration` の型追跡で `typeMap` にない型名は無視されていた。Enum 名は `env.getConst(effectiveType)` で plain object として取得できるため、`__vbaClass__` / `__vbaTypeName__` が両方なければ Enum 型として判定し `setVariableType({ vbaType: 'Long' })` を登録するよう修正。VBA では Enum の基底型は Long。レグレッションテスト: `tests/spec/enum.test.ts` Bug CB ブロック。 | このセッション |
| ~~**Bug CC: `Public Code As String * 5` クラスフィールドが固定長を無視していた（代入後も長さが保持されない）**~~ | **修正済み**: `createInstanceFromDef()` で `mt === 'string'` の初期化を `decl.fixedLength !== undefined ? ' '.repeat(fixedLength) : ''` に変更し、固定長フィールド名 → 長さのマップ `classFixedLengths` を構築してインスタンスに `__fixedLengths__` として付加。`evaluateAssignmentToVariable` の `__vbaClass__` 分岐（MemberExpression / ImplicitWithObjectExpression 両方）にUDT と同じ truncate/pad 処理を追加。レグレッションテスト: `tests/spec/class-module.test.ts` Bug CC ブロック。 | このセッション |
| ~~**Bug CF: `TypeOf obj Is InterfaceName` — `obj` のクラスが `Implements InterfaceName` していても `False` を返していた**~~ | **修正済み**: `evaluateTypeOfIsExpression` が `obj.__vbaTypeName__ === typeName` のみチェックしており、Implements チェーンを参照していなかった。`obj.__classDef__.body` を走査して `ImplementsDirective.interfaceName` と照合するブランチを追加。`TypeOf dogObj Is Animal`（Dog implements Animal）が正しく `True` を返すようになった。レグレッションテスト: `tests/spec/implements.test.ts` Bug CF ブロック。 | このセッション |
| ~~**Bug CE: `w.Inner("key")` — Property Set が Property Get より先に宣言されていると、読み取りコンテキストで Property Set が誤って呼ばれ Error 424 になっていた**~~ | **修正済み**: `evaluateCallExpression` の VBA クラスブランチで `classDef.procedures.find()` は宣言順で最初のものを返すため、Property Set が先なら setter が選択されていた。修正: `propertyType === 'get'` を優先する二段階 find に変更（Get が存在すれば Get を選択、なければ最初のマッチを選択）。読み取りコンテキストでは常に getter が優先されるべきであり、setter への呼び出しは代入文（`x = w.Inner` でなく `w.Inner = x`）から来るため、この変更は代入パスに影響しない。レグレッションテスト: `tests/spec/class-module.test.ts` Bug CE ブロック。 | このセッション |
| ~~**Bug CD: `Public Next As Object` — VBA キーワード `Next` がクラスフィールド名として受け付けられなかった**~~ | **修正済み**: `parser.ts` の `CONTEXTUAL_KW_STRUCTURAL` に `TokenType.KeywordNext` を追加。`Next`（トークン型 7）は `KeywordBase`（82）より小さく `KeywordBase..KeywordAddressOf` 範囲に含まれないため `isIdentifier()` が false を返していた。class body field parsing（line 2331 の `isIdentifier(inner)`）と `parseDimDeclarationList`（line 1868 の `isIdentifier(idToken)` チェック）の両方で `Next` がフィールド名として受け入れられるようになった。For ループの `Next` 終端判定はトークン型を直接比較するため影響なし。レグレッションテスト: `tests/spec/class-module.test.ts` Bug CD ブロック（連結リスト走査 + `n.Next Is Nothing`）。 | このセッション |
| ~~**Bug CG: `Dim arr(n) As Variant : arr(0) = 42 : TypeName(arr(0))` が `"Double"` を返していた**~~ | **修正済み**: `evaluateAssignmentStatement` の数値サブタイプ追跡が `Identifier` 変数のみを対象にしており、`CallExpression`（配列要素添字）を無視していた。Variant 配列への数値代入時にインデックスキー（`"0"` / `"1,2"` 等）→サブタイプのマップ `arr.__vbaSubtypes__` を更新するブランチを追加。`resolveNumericSubtype` にも `CallExpression` ケースを追加し、`TypeName`/`VarType` 評価時に `arr.__vbaSubtypes__[idxs.join(',')]` を参照するよう変更。レグレッションテスト: `tests/spec/typename.test.ts` Bug CG ブロック。 | このセッション |
| ~~**Bug CI: 複数モジュール構成で `Dim obj As New ClassName` が `Long 0` に化けていた**~~ | **修正済み**: `evaluateVariableDeclaration` の Enum-typed 変数チェック（Bug CB 対処）が `VbaNamespaceRef`（モジュール名として env に登録されたセンチネル）を enum と誤判定し `setVariableType('Long')` を呼んでいた。`!(enumObj instanceof VbaNamespaceRef)` の条件を追加して除外。`evalVBAModules` でクラス名と同名のモジュールを別ソースとして読み込んだ場合も `Dim obj As New MyClass` が正しく auto-instance placeholder を生成し、メンバーアクセスが動作するようになった。レグレッションテスト: `tests/spec/class-module.test.ts` Bug CI ブロック・`tests/spec/parse-as-class.test.ts` 全テスト。 | このセッション |
| ~~**Bug CH: `Function F(s As String * 5)` に短い/長い文字列を渡しても `s` が切り詰め/パディングされなかった**~~ | **修正済み**: パラメーター束縛時の `setVariableType` 呼び出し（3か所: 外部 `callProcedure`・VBA 内 `callProcedure`・クラスメソッド `callClassMethod`）に `fixedLength: mapped === 'String' ? param.fixedLength : undefined` を追加。`setLocally` → `coerceToType` のパスで `fixedLength` がある場合に truncate/pad が適用されるようになった。`LenOf("Hi")` → 5（パディング後）/ `LenOf("TooLongStr")` → 5（切り捨て後）が正しく返るようになった。レグレッションテスト: `tests/spec/fixed-length-string.test.ts` Bug CH ブロック。 | このセッション |
| ~~**Bug CA: `With m / .Data(0,0) = 1` — `CallExpression` の callee が `ImplicitWithObjectExpression` のとき Error 5 になっていた**~~ | **修正済み**: `evaluateAssignmentToVariable` の `CallExpression` 分岐に `call.callee.type === 'ImplicitWithObjectExpression'` の処理がなかった。`With` ブロック内で `.ArrayField(i,j) = val` や `.Property(arg) = val` を書くと `else` ブランチの Error 5 に落ちていた。VBA クラスの配列フィールド・Property Let に対応する分岐を追加（`MemberExpression` ケースと同じ論理、`obj` は `withObjectStack` から取得）。レグレッションテスト: `tests/spec/class-module.test.ts` Bug CA ブロック（1D/2D 配列フィールド）。 | このセッション |
| ~~**Bug BZ: `c.TheObj("k")` — `Property Get TheObj() As Object` が無引数のとき引数を Property Get への呼び出し引数と誤解釈し Error 450 になっていた**~~ | **修正済み**: `evaluateCallExpression` の VBA クラス分岐で `proc.isProperty && proc.propertyType === 'get' && proc.parameters.length === 0 && expr.args.length > 0` の場合に Property Get を引数なしで呼び出し、返却値に対して引数を添字アクセス（Dictionary/Collection/配列/デフォルトプロパティ）する分岐を追加。`c.Dict("key")`（Dictionary）/ `c.Items(n)`（Collection）/ `c.Data(i)`（配列）すべて正しく動作するようになった。レグレッションテスト: `tests/spec/class-module.test.ts` Bug BZ ブロック。 | このセッション |
| ~~**Bug BY: `Public P As New ClassName` クラスフィールドが `vbaNothing` に初期化されていた（`c.P.X = 42` が Error 91）**~~ | **修正済み**: `createInstanceFromDef()` のフィールド初期化ループが `decl.isNew` フラグを無視していたため、`Public P As New Point` のようなフィールドが `vbaNothing` として初期化されていた。`decl.isNew && classDefinitions.has(mt)` の分岐を追加し、`Collection` / 外部 factory / ユーザー定義クラスを即時インスタンス化するよう修正。レグレッションテスト: `tests/spec/class-module.test.ts` Bug BY ブロック（直接チェーンアクセス `c.P.X` + ネスト With `.P` 両方）。 | このセッション |
| ~~**Bug BX: `ByVal` UDT パラメーターがコピーを作らず参照として渡されていた**~~ | **修正済み**: `evaluateCallExpression` のパラメーター束縛時、`param.isByVal` かつ引数値が UDT オブジェクト（`__vbaTypeName__` あり・`__vbaClass__` なし）の場合に新メソッド `deepCopyUdtValue()` でコピーを作るよう修正。`deepCopyUdtValue()` はネストした UDT メンバーも再帰コピーする。VBA クラスインスタンス（オブジェクト参照型）はコピーせず参照をそのまま渡す（VBA の値型/参照型の区別に準拠）。 | `50485d4` |
| ~~**Bug BW: パラメーター宣言 `As String * N` が ParseError になっていた**~~ | **修正済み**: `parseParameter()` が fixed-length-string-spec (`String * N`) を解析しておらず、`ByRef s As String * 5` のような宣言で「Expected ')' after procedure parameters」のパースエラーになっていた。変数宣言の `parseVariableDeclaration()` と同様に、`As` 後の型名が `"String"` で次トークンが `*` の場合にサイズを読み込むよう修正。`Parameter` インターフェースに `fixedLength?: number` を追加。 | `6c97cd7` |
| ~~**Bug BV: `Dim a As New ClassName` 後の `TypeOf a Is ClassName` が `False` を返していた**~~ | **修正済み**: `evaluateTypeOfIsExpression()` で `evaluateExpression()` の後に `resolveAutoInstance()` を呼ばずプレースホルダーを渡していた。auto-instance placeholder は `__vbaTypeName__` を持たないため常に `vbaFalse` を返していた。修正: TypeOf 評価の先頭で `resolveAutoInstance(expr.expression, raw)` を呼ぶよう変更。`TypeOf a Is Object` は元から動作していた（プレースホルダーも `typeof obj === 'object'` を満たすため）。 | `aea8a33` |
| ~~**Bug BT: `Type` メンバーに配列 `Items(9) As Item` を持つ UDT で `c.Items(0)` が Error 438 になっていた**~~ | **修正済み（複合修正）**: (1) パーサーの `TypeMember` インターフェースに `isArray/arrayBounds` を追加し、配列境界（`(n)` / `(m To n)`）をパース・保持するよう変更（従来はスキップして捨てていた）。(2) `instantiateType()` で配列メンバーを初期化するよう修正。(3) 評価器の `evaluateCallExpression` に UDT 配列メンバーの読み取り、`evaluateAssignmentToVariable` に UDT 配列メンバーへの書き込みパスを追加。`isStatementAmbiguous` の `=` チェックも `isAssignableTarget` に変更（`parts(1) = "a"` の回帰を修正）。 | `671edd3` |
| ~~**Bug BU: `"7" = 7` が `False` を返していた（文字列/数値混在の `=` / `<>` 比較）**~~ | **修正済み**: JS の `===` は型変換しないため `"7" === 7` は `false`。`<`/`>`/`<=`/`>=` は JS が暗黙変換するので機能していたが `=` / `<>` は動いていなかった。`evaluateBinaryExpression()` の `case '='` と `case '<>'` に `typeof leftVal === 'string' && typeof rightVal === 'number'`（逆も）の条件を追加し、`toVbaNumber()` で文字列を数値に変換してから比較するよう修正。副作用として `live-vars.ts` の `getStmtUses()` に `DebugPrintStatement` ケースを追加（`Debug.Print x` の `x` がデッドストア検出の「使用」として計上されない回帰を修正）。 | `b049cd4` |

### ~~未修正バグ（評価 #34 で発見）~~

| 問題 | 最小再現コード | 根本原因 |
|---|---|---|
| ~~**Bug 34-A: 2桁年の日付リテラルが常に 20xx になる**~~ | **修正済み**: `Format(#3/15/30#, "yyyy-mm-dd")` → `1930-03-15`、`#3/15/99#` → `1999-03-15`。`parseDateLiteral()` を DateSerial と同じ既定 2029 ルール（00–29 は 20xx、30–99 は 19xx）へ修正し、`tests/spec/date-literal-parsing.test.ts` に回帰テストを追加。 | 旧実装は `year < 100` のとき常に `year += 2000` としていた。 |

---

## 未テスト・未探索の領域（今後の評価で優先すること）

評価済みドメインでカバーしたものを除いた、まだ十分に試されていない機能。

### カバレッジ計測で特定した暗部（2026-07-18、`./scripts/coverage.sh` 全 913 プロセスのユニオン集計）

行カバレッジ: `evaluator.ts` 92.1% / `parser.ts` 92.6% / `builtins.ts` 94.8% / `coerce.ts` 93.8% / `lexer.ts` 98.4%。
テストスイートが一度も通していない主な分岐（= 監査もされていない挙動。今後の評価ドメイン選定・テスト追加の優先候補）:

- ~~**日付リテラルの 2 要素形式**~~ **評価 #36 で実機確認・回帰テスト追加済み**: 年なしは実行年を使い、`#3/15#` は mm/dd、月として不正な `#15/3#` は dd/mm にフォールバックし、時刻部分を保持する。
- ~~**`&H`/`&O` 文字列の数値強制**~~ **実装・差分コーパス登録済み**: `vbaToNumber` が `&H`/`&O` を16/8進数として扱い、`CDbl("&H10")=16`、`CLng("&O17")=15` を確認。
- ~~**`Lock` / `Unlock` 文**~~ **評価 #35 で修正済み**: ハンドルごとの範囲ロックを管理し、閉鎖済み番号・重複範囲・未ロック解除を VBA エラーにする。
- ~~**`Open` 文の Access/Lock 節**~~ **評価 #36 で実機確認・修正済み**: `Lock Read` は他ハンドルの Read、`Lock Write` は Write を拒否し、`Shared` は両方を許可する。競合は Error 70。
- ~~**AppActivate / SendKeys の文形式**~~ **評価 #39: `AppActivate "Microsoft Excel"` / `SendKeys "{ESC}", True` は構文・実行とも正常。GUI 副作用や `Wait` は安全なスタブとしてエミュレートしない。**
- ~~**財務関数の一部経路**~~ **評価 #44: IRR/MIRR の1基底配列を `NaN` から修正。IPmt/PPmt の第2期以降の符号、期首払い、および範囲外 period を修正。**
- ~~**Decimal 除算経路**~~ **評価 #41: `1/6` / `1/7` の28桁丸め、負値、`±1.005`・`±1.015` の銀行丸め、Collection 集計を確認。期待どおり。**
- ~~**`Set obj.Prop = x` のチェーン経路**~~ **評価 #40: クラスの `Property Set` と Dictionary の `Item` を組み合わせたキー付き左辺・右辺で、連鎖代入と参照保持が正常。**
- ~~**`DateAdd "ww"`・`DatePart` の年始基準系**~~ **評価 #42: `DateAdd("ww")` は正常。`DatePart("ww", #1/1/2021#, vbSunday, vbFirstJan1 / vbFirstFourDays / vbFirstFullWeek)` を `1,1,1` から `1,53,52` へ修正。**
- ~~**クラス本体直下の `Dim`/`Static` 宣言**~~ **評価 #43: `Dim` フィールド・Collection と `Static hits As Long` を評価。同一インスタンスで状態を保持し、別インスタンスは独立する。**
- ~~**パーサーのエラー回復**~~ **評価 #45: `errorRecovery: true` では不正手続きの診断後に後続手続きを保持する。通常実行時は例外を返す設計。**
- ~~**`Declare` 文の一部形式**~~ **評価 #46: `Declare PtrSafe` の `Lib` / `Alias` / `LongPtr` / `ByRef String` / 戻り型を含む有効な宣言はロード可能。外部 DLL 呼び出しは安全なスタブ。**

### 数学・数値関数・型チェック・文字列変換（評価 #22 で確認済み）

- ~~`Sqr()`, `Abs()`, `Log()`, `Exp()`, `Sin()`, `Cos()`, `Atn()`, `Sgn()`, `Fix()`, `Rnd()`/`Randomize`~~ **評価済み（評価#22）: 全正常動作。`Atn(1)*4` で Pi 計算も正常。`Fix(-2.7)=-2` / `Int(-2.7)=-3` の違い正確。`Randomize seed` で同一シード → 同一乱数列を確認。**
- ~~`TypeName()`, `VarType()`, `IsNull()`, `IsArray()`, `IsObject()`, `IsError()`, `IsEmpty()`~~ **評価済み（評価#22）: `TypeName(42)="Integer"` など VBA リテラル引数は正常。`IsNull`/`IsEmpty`/`IsArray`/`IsObject`/`IsError`/`CVErr` すべて正常。ただし **Bug 22-1**: Variant 変数経由の TypeName/VarType が subtype を失い常に Double を返す。**
- ~~`Like` 演算子~~ **評価済み（評価#22）: `*`/`?`/`[a-z]`/`[!x]`/`#` すべて正常動作。メールバリデーションパターン `"*@*.*"` も正常。**
- ~~`Optional ByVal x As Long = 0` デフォルト値付き Optional パラメーター~~ **評価済み（評価#22）: 正常動作。部分指定（一部の Optional のみ渡す）も正常。**
- ~~`ParamArray args()` 可変引数~~ **評価済み（評価#22）: `CalcMean(ParamArray vals())` / `CalcStdDev(ParamArray vals())` 正常動作。`eval()` から `CalcMean(1,2,3,4,5)` で呼び出し可能。**
- ~~名前付き引数 `Func(argB:=2, argA:=1)`~~ **評価済み（評価#22）: `FormatNum(value:=3.14, decimals:=3, prefix:="PI=")` 正常。引数順序逆転（`decimals:=4, value:=2.71`）も正常。部分指定（`value:=99.9` のみ）も正常。**
- ~~`IIf()`, `Choose()`, `Switch()` 条件関数~~ **評価済み（評価#22）: 全正常。`Choose(dayNum,"Mon","Tue",...)` 1-based index / `Switch(score>=90,"A", True,"F")` フォールバック True も正常。**
- ~~`Null` の扱い: `IsNull()`, `Null` リテラル, `Null` と `Empty` の区別~~ **評価済み（評価#22）: `IsNull(Null)=True` / `IsEmpty(Null)=False` / `Null+1=Null`（伝播）/ `TypeName(Null)="Null"` / `VarType(Null)=1` すべて正常。VarType 0=Empty, 1=Null の区別も正確。**
- ~~`Asc()`, `Chr()`, `Hex()`, `Oct()`, `StrReverse()`, `Space()`, `String()`, `StrConv()`~~ **評価済み（評価#22）: 全正常。`Hex(65)="41"`, `Oct(65)="101"`, `StrReverse("Hello")="olleH"`, `String(5,"-")="-----"`, `Space(3)="   "`, `StrConv("hello",1)="HELLO"` / `StrConv("HELLO",2)="hello"` すべて正常。**

### ファイル入出力

- ~~`Open ... For Output/Input/Append`, `Print #`, `Line Input #`, `Close`~~ **評価済み（評価#5）**
- ~~Sandbox パス変換（`C:\` → サブディレクトリ変換）の動作確認~~ **評価済み（評価#5）: 正常動作**
- ~~存在しないファイルの `Open For Input` → Error 53 (File not found)~~ **評価済み（評価#5）: 正常動作**
- ~~`Open For Binary` / `Put #` / `Get #` / `Seek`~~ **評価済み（評価#36）: `Byte`/`Integer`/`Long`、CP932文字列、配列なしUDTのバイナリ入出力を実機差分で確認。配列・可変長文字列を含むUDT、Date/Currency は未対応。**
- ~~`Open For Random As #n Len = recLen`~~ **評価済み（評価#26）: `Len =` 節がパースエラー（Bug 26-3）。未実装**
- ~~`FileLen(path)` / `FileDateTime(path)` / `Kill path`~~ **評価済み（評価#26）: 全正常動作**
- ~~`GetAttr(path)`~~ **評価済み（評価#26）: Error 35 未実装（Bug 26-7）**
- ~~`Error(n)` 関数~~ **評価済み（評価#26）: 主要コード（5/6/9/11/13/53/91）は正確。未登録は汎用フォールバック**

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
- ~~`CDbl` / `CSng` の精度と `Single` 型の動作~~ **評価済み（評価 #21）: `CSng(1)/CSng(3)` → `0.3333333432674408`（32bit 精度）、`CDbl(1)/CDbl(3)` → `0.3333333333333333`（64bit 精度）。実 VBA と同様の精度差を正しく再現。正常動作。**

### コレクション

- ~~VBA `Collection` オブジェクト（`Add` / `Item` / `Count` / `Remove` / `For Each`）~~ **評価済み（評価#4）**
- ~~ゼロ基底 vs 1基底のインデックスの挙動~~ **評価済み（評価#4）: 1基底が正しく動作する**

### 複数クラスの連携

- ~~クラス間の相互参照（A が B のインスタンスを持つ）~~ **評価済み（評価#7）: 正常動作**
- ~~`Set` 代入 / `Is Nothing` 判定~~ **評価済み（評価#7）: 正常動作**
- ~~`Class_Terminate` の呼ばれるタイミング~~ **評価済み（評価#7）＋修正済み: 参照カウント実装により最後の参照が解放されたときに発動するよう修正**

### 拡張機能 LSP（未テスト・要評価）

- ~~`__mocks__` ディレクトリによる VBA/JS/TS モック注入~~ **評価済み（評価 #12）: JS `__addCreateObject__` / VBA `.cls` クラスモック / `__mocks__.bas` 単一形式 / `.ts` モック すべて正常動作。同名 Public Function を複数 VBA モックが持つ場合は Ambiguous procedure バグあり**
- ~~`setBuiltinOverride` で組み込み関数を上書き~~ **評価済み（評価 #12）: MsgBox を vbOK/vbCancel に固定 → 正常動作。`spy()` API も動作確認**
- ~~vba-types.json の自動リロード（FileSystemWatcher）~~ **評価済み（評価 #12 ソース確認のみ）: `extension.ts:70-77` に `createFileSystemWatcher` 実装あり。`onDidCreate` / `onDidChange` でリロード。`onDidDelete` は未実装（型スタブが残存する小欠陥）**
- ~~Formatter（コード整形）~~ **評価済み（評価 #14）: キーワード大文字小文字 / インデント / 識別子ケーシング 全正常。組み込み型名（String/Integer 等）は正規化しない（設計ギャップ）**
- ~~CodeLensProvider（▶ Run / 🐛 Debug / 参照数 / テスト状態）~~ **評価済み（評価 #14）: 基本動作正常。`Test_*` が常に `✓ Tested`（疑似陽性バグ）**
- ~~FoldingRangeProvider（コード折りたたみ）~~ **評価済み（評価 #14）: 全主要ブロック対応・正常動作**
- ~~CallGraphProvider（コールグラフ構築）~~ **評価済み（評価 #14）: マルチファイル・相互再帰・自己再帰・Excel依存検出すべて正常**
- ~~With ブロック内のチェーン補完（`.Cells(args).` など）~~ **評価済み（評価 #10, #11）: 引数なし `.Cells.` は正常動作、引数付き `.Cells(1,1).` は評価 #10 でバグ検出 → コミット `3300dcb` で修正 → 評価 #11 で修正確認（48 件返るようになった）**
- ~~クロスモジュール補完（複数ファイル展開時）~~ **評価済み（評価 #10）: `parseAsClass` オプションを使えば動作する。単純な statements マージだけでは不十分**
- ~~`generateDefaultTypeStubsJson` / `setTypeStubs`~~ **評価済み（評価 #10, #11）: コミット `9e41d0c` で `parseTypeStubsJson` 追加により API 非対称が解消。`generateDefaultTypeStubsJson()` → `parseTypeStubsJson(json)` → `cp.setTypeStubs(map)` の 3 ステップで完結**
- ~~カスタム型上書き優先~~ **評価済み（評価 #10）: BUILTIN_MEMBERS よりカスタム定義が優先される。正常動作**
- ~~VBA016 Quick Fix の動作（initTypeStubs / addToTypeStubs コマンド）~~ **評価済み（評価 #11）: `extension.ts` に完全実装。`vba-runner.initTypeStubs`（デフォルト vba-types.json 生成）と `vba-runner.addToTypeStubs`（型名追記）が CodeAction として登録済み。VS Code なしの動作確認は不可**
- ~~シグネチャヘルプの深掘り（ネスト呼び出し・Optional 引数・ParamArray）~~ **評価済み（評価 #11）: 組み込み 60 件超・ユーザー定義 Function・ネスト優先すべて正常。文字列リテラル入力中に消えるバグあり**
- ~~DefinitionProvider（F12 定義ジャンプ）~~ **評価済み（評価 #13）: `.bas` は正常。`.cls` は `parseAsClass` 使用時に `loc` 未設定バグにより常に `null`（Bug B）**
- ~~ReferencesProvider（Find All References）~~ **評価済み（評価 #13）: `.bas` は概ね正常。`FuncName = value` 戻り値代入行が refs に含まれる（仕様準拠）。ローカル変数スコープ絞り込み正常。`.cls` は Bug B によりスコープ・includeDeclaration 除外が無効**
- ~~RenameProvider（F2 リネーム）~~ **評価済み（評価 #13）: `.bas` 正常（戻り値代入行も含む TextEdit は正しい）。`.cls` は Bug B によりスコープ無視のテキスト置換**
- ~~SymbolProvider（ドキュメントシンボル / アウトライン）~~ **評価済み（評価 #13）: `.bas` 完全動作。`source` 引数なし→Sub/Function のみ、あり→セクションヘッダー `' --- Name ---` も SymbolKind.Namespace で追加。`.cls` は動作するが全シンボルの loc が `(0,0)`（Bug B の影響・fallback 動作）**
- `formatter.ts` — VBA コード整形（インデント・キーワード大文字化・空行）**（評価 #14 進行中）**
- `code-lens-provider.ts` — ▶ Run / ▶ Test CodeLens 生成**（評価 #14 進行中）**
- `folding-range-provider.ts` — Sub/If/For 等の折りたたみ範囲**（評価 #14 進行中）**
- `call-graph-provider.ts` — コールグラフ（呼び出し関係・再帰検出）**（評価 #14 進行中）**
- ~~`auto-parens.ts`~~ **評価済み（評価 #15）: `autoParensEdit`/`getBlockEnd`/`needsBodyIndent`/`needsEndBlock` 正常動作。戻り型付き Function 未検出バグあり**
- ~~`keyword-casing.ts`~~ **評価済み（評価 #15）: `canonicalKeyword`/`isInStringOrComment` 正常動作**
- ~~`label-navigator.ts`~~ **評価済み（評価 #15）: GoTo←→LabelStatement 双方向ナビゲーション正常動作**
- ~~`line-continuation-checker.ts`~~ **評価済み（評価 #15）: `needsLineContinuation`/`stripInlineComment` 正常動作**
- ~~`variant-type-inferencer.ts`~~ **評価済み（評価 #15・#17）: `inferVariantTypes`・`inferProcedureHints`・`inferModuleVarType`・`inferLocalVarType`・`buildProcMap`・`findProcAtLine` 全関数確認。CreateObject ProgID型推論（Dictionary/FSO/Worksheet等）・曖昧型抑制・パラメーター Variant ヒント・戻り型ヒント・モジュールレベル変数推論すべて正常。`Dim x As Variant` + 数値代入 → Long ヒントが出る（設計通り）**
- ~~`test-discovery.ts`~~ **評価済み（評価 #15）: `TestDiscovery.discoverTests` 正常動作**
- ~~`test-runner.ts`~~ **評価済み（評価 #15）: `runTests()` はスタブ（常に passed バグ）。`runTestWithEvaluation()` は実装済みだがエラーメッセージが `"[object Object]"` になるバグ**
- ~~`ast-comparison.ts`~~ **評価済み（評価 #15）: `astEqual`/`serializeAst`/`findMatchingExpressions` 正常動作**
- ~~`hover-provider.ts`~~ **評価済み（評価 #15・#17）: `HoverProvider.getHoverInfo` ローカル変数・パラメーター・モジュールレベル変数・定数・プロシージャ・As Object→CreateObject型推論表示・到達定義情報 全正常動作。`Const MAX As Long = 100` のホバーが `Const MAX` のみ（型・値なし）バグあり。`Public`/`Private` が lowercase で表示されるバグあり**
- デバッガー系（VS Code なしでは動作確認不可・ソース確認のみ）: `debugger.ts` / `debug-adapter.ts` / `debug-session.ts` / `debug-worker.ts` / `vscode-debug-adapter.ts`
- `call-graph-webview.ts` — コールグラフの WebView 表示（VS Code なしでは動作確認不可）

### 使い勝手の観点

機能の正常動作だけでなく、ユーザー体験として自然かどうかを評価する観点。
サブエージェントは以下の視点で「引っかかり」や「改善提案」を記録すること。

- **ホバー表示**: 変数・型名・プロシージャのホバーに表示される情報が十分か、冗長・不足はないか
- **補完候補の順序と品質**: 最も使いそうな候補が上位に来るか、不要なノイズが多くないか
- **エラーメッセージの分かりやすさ**: VBA016 等の診断メッセージがユーザーにとって意味が伝わるか
- **初回ロード体験**: ファイルを開いた直後の診断・補完の応答速度や順序感（先にスキャンされるべきかどうか）
- **API 直感性（VBARunner）**: `run()` / `eval()` の引数・戻り値がドキュメントなしでも推測できるか
- **テスト結果の読みやすさ**: `runTests()` の出力（状態・メッセージ・時間）が CI ログとして見やすいか

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
13. ~~**`Currency` は固定小数点ではなく浮動小数点**~~（**評価 #23 で修正確認済み**）: `CCur(0.1) + CCur(0.2)` は厳密に `0.3`（内部 BigInt で 4桁固定小数点演算）。`VbaRunner.run()` / `.eval()` の戻り値も `number` に正規化される（Bug C-1/C-2 修正済み）。
14. **`empty` / `Empty` は変数名に使えない**: `Dim empty() As String` は「Expected variable name (Found empty)」でパースエラー。`emptyArr` 等の代替名を使う。
15. **`run()` の第2引数は省略可能**（修正済み）: 引数なし Sub は `r.run('Sub名')` と省略できる。`args` のデフォルト値を `[]` に修正済み。
16. ~~**時刻のみリテラル `#HH:MM:SS#` は使えない**~~（修正済み `b4d00c3`）: `#12:30:45#` / `#8:30:00 AM#` が正常動作するようになった。評価 #27 で動作を確認。日付+時刻 `#2024/01/15 12:30:45#` も正常。
87. ~~**`CallByName` の `VbLet(4)` / `VbSet(8)` は未実装（Bug 27-1）**~~: **修正済み**。`builtins.ts` の `callbyname` に VbLet(4)/VbSet(8) ブランチを追加。VBA クラスの `classDef.procedures` から `propertyType === 'let'`/`'set'` の Property を検索して `ctx.callMethod` で呼び出す。レグレッションテスト: `tests/spec/callbyname.test.ts`（4テスト）。
88. **`On Error GoTo` ハンドラー内で `Exit Function/Sub` 後も `Err.Number` が残留する（仕様準拠）**: エラーハンドラーで処理後 `Exit Function` で抜けると、呼び出し元の `Err.Number` にエラー番号が残留する。これは実 VBA と同じ動作。VBA の `Err` はグローバルオブジェクトであり、`Resume`/`Resume Next`/`Err.Clear`/`On Error` 文の実行でのみリセットされる。ワークアラウンド: エラーハンドラー内で明示的に `Err.Clear` を呼ぶ。
89. **`GetAllSettings` の戻り値は `Array[row][0=key, 1=val]` の2次元配列**: 評価 #27 で確認。`GetAllSettings(appName, section)` は行数×2列の配列を返す。TypeScript 側でアクセスするには `(result as any[][])[row][0]`（キー）/ `[1]`（値）でアクセスする。README に記載なし。
90. **`Shell(cmd)` はスタブとしてタスク ID `1` を返す**: 実際にコマンドを実行せず `[SHELL] cmd` のログを出力してタスク ID `1` を返す。CLAUDE.md に記載の通り。`DoEvents` は `0` を返す no-op。
91. **`Filter(arr, keyword, include, compare)` の第3引数 `include` と第4引数 `compare` は動作する**: `Filter(arr, "key", False)` で非一致フィルタが正常動作。`Filter(arr, "key", True, 1)` で `vbTextCompare`（大文字小文字無視）も動作。一致なしの場合は `UBound=-1` の空配列を返す。
41. ~~**`Dim s As String * N`（固定長文字列）は未実装**（Bug 21-1）~~: **修正済み**。`Dim s As String * 10` が正常に動作する。短い文字列はスペースでパディング、長い文字列は切り捨て、`Len(s)` は常に N を返す。UDT の `Name As String * 30` メンバーも対応。初期値は NUL 文字 × N（VBA 仕様準拠）。
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
32. **`getCompletions` に渡すのは `parse().body`（`Statement[]`）であり `Program` オブジェクトではない**: `parse()` は `{ type: 'Program', body: Statement[], diagnostics: [] }` を返す。`getCompletions(parsed, ...)` と渡すと `statements is not iterable` エラー。`getCompletions(parsed.body, ...)` が正しい。（ただしサブエージェントは `stmts` を `parse()` の返り値から `.body` で取り出して正常動作した）
33. **`parseTypeStubsJson` 関数を使えば `setTypeStubs` の前処理が不要になった（評価 #11 で確認）**: コミット `9e41d0c` で `parseTypeStubsJson` が `completion-provider.ts` に追加された。`generateDefaultTypeStubsJson()` → `parseTypeStubsJson(json)` → `cp.setTypeStubs(map)` の 3 ステップで完結。旧評価 #10 で必要だった手動 JSON.parse + lowercase 変換 + Map 構築は不要。
34. **`SignatureHelpProvider.getSignatureHelp(stmts, source, line, char)` が使える**: 組み込み関数 60 件超・ユーザー定義 Function・ネスト呼び出し（内側を優先）が動作する。ただし文字列リテラルを入力し始めると `findCallContext` が `null` を返してシグネチャが消える（`signature-help-provider.ts:145` の右→左スキャンバグ）。
26. **`CompletionProvider.getCompletions` の引数順は `(statements, source, line, character)`**: `statements` が第1引数、`source` が第2引数の順。ドキュメントや過去の記述で逆順に書かれていた例があるため注意。
27. **`checkUnknownTypes` の第2引数は `Set<string>`（オブジェクトリテラル `{}` は不可）**: 誤って `{}` を渡すと `knownTypeNames.has is not a function` で実行時エラー。`new Set<string>()` か `collectUserDefinedTypeNames(stmts)` の戻り値を渡すこと。
28. **~~`setTypeStubs()` の引数は `Map<string, CompletionItem[]>`（JSON 文字列ではない）~~（評価 #11 で解消確認）**: コミット `9e41d0c` で `parseTypeStubsJson(json)` が追加された。`generateDefaultTypeStubsJson()` → `parseTypeStubsJson(json)` → `cp.setTypeStubs(map)` の 3 ステップで完結。手動変換は不要。
29. **`parse()` フリー関数は存在しない**: `parser.ts` はフリー関数の `parse()` をエクスポートしていない。`Lexer` でトークナイズ後に `new Parser(tokens, { errorRecovery: true }).parse()` という形で呼ぶこと。
30. **クロスモジュール補完にはクラスモジュールのパースに `parseAsClass` オプションが必要**: `.cls` ファイルを `new Parser(tokens, { parseAsClass: 'ClassName' })` で解析しないと `ClassDeclaration` として AST に格納されない。単純に `statements` をマージするだけでは不十分で、クラスのメンバーが補完候補に出ない。
31. **~~With ブロック内の引数付きメソッドチェーン後の補完は未対応（新バグ）~~（コミット `3300dcb` で修正・評価 #11 で確認）**: `    .Cells(1, 1).` でトリガーして 48 件の Range メンバーが返るようになった。
35. **JS `__addCreateObject__` モックは `CreateObject(progId)` 呼び出し全体を置換する**: `factory()` が返すオブジェクトに `__progId__` がない場合は `mock-loader.ts:181` でキーを `__progId__` として補完する。`evaluator.registerComObject` に渡される。既存の組み込みスタブ（Scripting.Dictionary など）よりも優先される。
36. **`.ts` モックは `tsx` 環境下では `createRequire` で直接 require できる**: `mock-loader.ts` は `createRequire(import.meta.url)` で生成した `_require` を使って `.ts` モックを読み込む。`tsx` 実行環境では `.ts` ファイルを直接 require できるため、TypeScript でモックを書ける。
37. **複数の VBA モックが同名 `Public Function` を持つと Ambiguous procedure エラー**（評価 #12 で発見・未修正）: `__mocks__/A.bas` と `__mocks__.bas` が同名関数を定義すると `Ambiguous procedure` エラー。`mock-loader.ts` のドキュメントは "後勝ち" と書いているが、これは JS/TS モック（`setBuiltinOverride` で `env.set` 上書き）にのみ適用される。VBA モックは別モジュールとして `procedures` マップに追記されるため曖昧さエラーになる。ワークアラウンド: 同名関数を持つ VBA モックは1ファイルにまとめること。
38. **`__mocks__/ClassName.cls` は本番 `ClassName.cls` を上書きする**: `promoteMockVbaClasses` が mock の `.cls` を `externalObjectFactories` に昇格させ、`instantiateClass` で `classDefinitions` より優先されるため、本番クラスは実質上書きされる。VBA クラスモックは期待通り動作する。
39. **`vba-types.json` の FileSystemWatcher は削除（`onDidDelete`）を監視しない**: `extension.ts:70-77` では `onDidCreate` と `onDidChange` のみフックしている。`vba-types.json` を削除した場合は型スタブがメモリに残ったまま（拡張機能再起動まで消えない）。
40. **LSP ナビゲーション 4 プロバイダーの API パターン**（評価 #13）: 全プロバイダーとも `setDocumentUri(uri)` + 1 メソッド呼び出しのシンプルな設計。`stmts` には `new Parser(tokens, { errorRecovery: true }).parse().body` を渡す（`Program` オブジェクトをそのまま渡すと動かない）。`DefinitionProvider.getDefinition(stmts, source, line, char)` / `ReferencesProvider.getReferences(stmts, source, line, char, includeDeclaration)` / `RenameProvider.getRename(stmts, source, line, char, newName)` / `SymbolProvider.extractSymbols(stmts, source?)` の 4 メソッド。`line` / `character` はすべて 0-based。
41. **`SymbolProvider.extractSymbols(stmts, source)` は `source` あり時にセクションヘッダーを抽出する**（評価 #13）: `' --- Section Name ---` / `' === Section Name ===` スタイルの VBA コメントを `SymbolKind.Namespace` として返す。`.bas` の機能として有用。`source` 省略時は Sub/Function/Property のみ。
42. ~~**`parseAsClass` 使用時は DefinitionProvider / ReferencesProvider / RenameProvider が `.cls` シンボルを認識しない**（評価 #13・Bug B・修正済み）~~: 修正: `parseClassBody()` 内で `tok`（ブランチ開始前トークン）を startTok として記録し、各パース呼び出し後に `stmt.loc = { start: tok.line/col, end: endTok.line/col }` を設定。`ClassDeclaration` 本体の `loc` も最初・最後の有効トークンから設定。これにより `DefinitionProvider.getDefinition` が `.cls` ファイルで正しく位置を返す・`SymbolProvider` 子シンボルの位置が正確になる。`tsc -b` / `class-module.test.ts` 全通過確認済み。
43. **`ReferencesProvider` は `FuncName = value` の戻り値代入行を参照として返す**（評価 #13）: テキスト正規表現による全出現検索の仕様。`includeDeclaration: false` は宣言行の正確な position のみを除外する。戻り値代入 `CalcTax = price * rate` は宣言行 (line 0) とは別の行 (line 1) のため除外されない。**RenameProvider の観点では正しい動作**（関数リネーム時に戻り値代入も書き換え必須）。「Find All References」の「呼び出し一覧」として使う場合はユーザーが驚く可能性あり。
44. **`format(source, options)` + `applyEdits(source, edits)` の使い方**（評価 #14）: `format()` は `TextEdit[]` を返すだけ。実際に文字列を得るには `applyEdits(source, format(source, opts))` とする。`FormatterOptions` は `indentSize`（デフォルト 4）/ `indentChar`（デフォルト `' '`）/ `keywordCase`（`'pascal'` or `false`）。
45. **Formatter が正規化しないもの**（評価 #14）: 組み込み型名（`String`/`Integer`/`Boolean`/`Long`/`Double`/`Date`/`Variant`/`Object` など `As` 後に使う型名）と組み込みオブジェクト（`Debug`/`Err`/`ActiveSheet` など）は `TokenType.Identifier` として扱われ、ユーザー宣言もないため正規化されない。`Dim s As string` → `Dim s As string`（変化なし）。
46. **Formatter の行継続インデント方針**（評価 #14）: 行継続（`_` 終端）行の後続行は `actualIndent`（既存の先頭空白）を保持するだけで、新たなインデントを付与しない。継続行が列 0 に書かれていれば列 0 のまま。意図的保存動作（フォーマッター非制御領域）。
47. **`CodeLensProvider.getCodeLens` の lens 構成**（評価 #14）: プロシージャごとに最大 5 種類の lens が付く: `▶ Run` + `🐛 Debug`（必須パラメーターなし or `Test_*` 1param のみ）/ 参照数 or `🔔 Event Handler`（常時）/ `✓ Tested` or `Untested`（常時）/ `📊 Show in Call Graph`（常時）。テスト結果を渡した場合は `Test_*` にさらに結果 lens（`✓ Xms` or `✗ message`）が追加。
48. **`CodeLensProvider.getDeadCodeWarnings` で Private 0参照プロシージャを検出できる**（評価 #14）: `isPrivate && refCount === 0 && !isEventHandler` の条件でデッド候補を列挙。イベントハンドラー（`KNOWN_VBA_EVENT_NAMES` 照合）は除外される。
49. **`FoldingRangeProvider.getFoldingRanges` は全主要ブロックをカバー**（評価 #14）: Sub/Function/If/For/ForEach/DoWhile/While/With/SelectCase/Type/Enum/Class が対象。単行 `If x Then Debug.Print y`（Then 後に実体がある場合）は折りたたみ対象外。`FoldingRange` は `{ startLine, endLine }` の 0-based ペア。
51. ~~**`autoParensEdit` は戻り型なしの Sub/Function/Property のみ対応**（評価 #15）~~（評価 #16 で修正確認済み）: `PROC_NO_PARENS` 正規表現に `(?:\s+As\s+\w+)?` が追加され、`Function GetValue As Long` → `{ insertCol: 17 }` が正常に返る。
52. ~~**`TestRunner.runTests(statements)` はスタブ実装**（評価 #15）~~（評価 #16 で修正確認済み）: `runTests` の引数は `src: string`（VBA ソース文字列）に変更され、内部で実際に評価を行う本実装になっている。`Err.Raise` するテストは `{ state: 'failed' }` を返す。
53. ~~**`TestRunner.runTestWithEvaluation` のエラーメッセージが `"[object Object]"` になる**（評価 #15）~~（評価 #16 で修正確認済み）: `catch` 節が `(testError as any)?.message ?? String(testError)` になっており、VBA エラー plain object の `.message` が正しく取り出される。
56. **`TestRunner.runTests(src)` の正しい引数は VBA ソース文字列**（評価 #16）: 評価 #15 では `stmts`（Statement[]）が引数と記録されていたが、現行は `src: string` が正しい。`runTests(parsed.body)` のように Statement[] を渡すと `char.codePointAt is not a function` で TypeError になる。正しくは `runTests(vbaSourceCode)` とすること。
54. **`HoverProvider.getHoverInfo` のパラメーター hover range は列 0 固定**（評価 #15）: パーサーがパラメーターの loc を AST に記録しないため、`symbol-table.ts` がパラメーターの位置を手続きヘッダー行・列 0 からの推定値として記録する。ホバー時の下線範囲が列 0 から始まる（設計上の既知の制限）。
55. **`inferVariantTypes` は `buildProcMap` の返す `Map` を渡すと再帰型推論が効く**（評価 #15）: `allProcs: Map<string, ProcedureDeclaration>` を渡さないと関数呼び出し経由の型推論が行われない。同ファイルの全手続きマップを作って渡すこと。同じ型推論を繰り返す場合は `memo: Map<string, InferredType>` を共有すれば高速化できる。

50. **`CallGraphProvider.buildCallGraph` はマルチファイル対応・再帰検出あり**（評価 #14）: `fileMap: Map<uri, { statements, uri }>` を渡す。戻り値は `{ nodes: Map<nameLower, ProcNode>, edges: CallEdge[] }`。自己再帰（`factorial → factorial`）・相互再帰（`funca ↔ funcb`）ともにエッジとして正しく検出。メンバーアクセス呼び出し（`obj.Method()`）は追跡しない（設計上の仕様）。`ProcNode.isExcelDependent` は EXCEL_ROOT_OBJECTS 定数セット（`sheets`/`range`/`cells`/`application` 等）への直接参照で判定。
57. **`inferProcedureHints` は `inferVariantTypes`（変数）+ パラメーターヒント + 戻り型ヒントの上位 API**（評価 #17）: 変数・パラメーター・戻り型の3種類のヒントを一括取得するには `inferProcedureHints` を使う。`inferVariantTypes` は変数のみ。戻り型ヒントは `kind: 'return'` で区別できる。
58. **`Dim x As Variant` に数値を代入すると `As Long` ヒントが出る（設計通り）**（評価 #17）: `variant` 宣言済みでも代入から推論された型が異なれば（'long' ≠ 'variant'）ヒントが表示される。ユーザーにとっては「Variant で宣言しているのにヒントが出る」と驚く可能性がある。意図的な動作。
59. **`lintProgram(program)` に渡すのは `Program`（`parse()` の戻り値そのもの）**（評価 #17）: `checkUnknownTypes` と異なり `lintProgram` は `program.body` ではなく `Program` オブジェクト全体を受け取る。`Program.body` を渡すと `checkOptionExplicit` が機能しない（`OptionExplicitStatement` を走査できないため）。
60. **診断コード VBA001〜VBA014 の severity 一覧**（評価 #17）: VBA013=Error(1)、VBA001/VBA005/VBA006/VBA008/VBA009/VBA010/VBA012/VBA014=Warning(2)、VBA002/VBA004/VBA007=Info(3)、VBA003=Hint(4)。VBA011（Range変数アクセス）と VBA016（未知型・別モジュール）は `lintProgram` には含まれない（`checkUnknownTypes` で別途取得）。
61. ~~**`Const` の hover displayText にはすべての情報が含まれていない**（評価 #17 で発見）~~: **修正済み（`03f2b23`）**: `parser.ts` で `ConstDeclaration.objectType` を保持し、`symbol-table.ts` に `constLiteralText()` ヘルパーを追加。`Const MAX As Long = 100` と型・値つきで表示されるようになった。
62. ~~**`Public`/`Private` が hover で lowercase 表示される**（評価 #17 で発見）~~: **修正済み（`03f2b23`）**: `symbol-table.ts` に `cap()` ヘルパーを追加し、scope 文字列を capitalize するよう修正。
63. **`analyzeDefUse(proc, startLine, endLine)` の行番号は 1-based**（評価 #18）: `src/engine/def-use-analyzer.ts` の公開 API。`LspServer.getCodeActions` 内部では 0-based の `range.start.line` を `+1` して渡している。呼び出し元で要注意。戻り値 `{ inputs, outputs, locals }` はすべて小文字正規化済み変数名の配列。
64. **`findDeadStores(proc)` は `ProcedureDeclaration` を直接受け取る**（評価 #18）: `src/engine/dead-store.ts`。`DeadStore` の `line`/`column`/`endColumn` はすべて 1-based で診断 range に直接使える。ByRef パラメーターと Function 戻り値変数は `alwaysLive` 扱いで除外される（正しい動作）。
65. **`LspServer.getCodeActions` は `analyzeDefUse` の実行まで完了しており、`buildExtractFunctionEdit` は純粋テキスト変換のみ**（評価 #18）: `getCodeActions` の戻り値 `actions[0].command.arguments` には `[uri, range, defUseResult, procSignature, callStatement]` が含まれる。ユーザーが procName を指定してから `buildExtractFunctionEdit(uri, range, procName, defUseResult, procSignature, callStatement)` を呼ぶ2ステップ設計。VS Code コマンドとして仲介される想定。
66. ~~**`buildExtractFunctionEdit` が選択内の Dim をパラメーターと共存させる不正コードを生成する（Bug R1）**（評価 #18 で発見）~~: **修正済み（評価 #18）**: `reindented` 生成時に inputs/outputs に含まれる変数名の `Dim` 行をフィルタリングする処理を追加（`server.ts:819-830`）。レグレッションテスト: `tests/lsp/lsp-code-actions.test.ts` Test 13。
67. **`buildExtractFunctionEdit` のパラメーター名は小文字正規化される**（評価 #19）: `analyzeDefUse` が返す変数名配列が小文字正規化済みのため、シグネチャの `ByVal inputVal As Variant` → `ByVal inputval As Variant` になる。VBA は大文字小文字無視のため動作には影響しないが可読性が低下する。本体コードの識別子は元ソースの大文字小文字を維持するため、シグネチャと本体で不一致になる。
68. **`buildExtractFunctionEdit` は抽出先を常に `Private Sub` にする**（評価 #19）: 元のコードが Function 内であっても、戻り型の推論が困難なため抽出先は常に `Private Sub` になる。設計上の制限。
69. **`findDeadStores` はループカウンター・集計変数・条件分岐内代入・関数戻り値変数を誤検出しない**（評価 #19）: CFG ベースの生変数解析が高精度で動作することを確認。ByRef パラメーターと Function 戻り値変数は `alwaysLive` 扱いで適切に除外。
70. ~~**With ブロック内でユーザー定義クラスの Property Get が Error 424（Bug W1）**（評価 #19 で発見）~~: **修正済み（評価 #19）**: `evaluateImplicitWithObjectExpression` に `__vbaClass__` ブランチを追加。`With classInstance: .PropertyGet` が正常動作するようになった。`class-module.test.ts` Bug W1 テスト追加。
71. ~~**Variant 変数経由の TypeName/VarType は subtype を失い常に Double を返す（Bug 22-1）**（評価 #22）~~: **修正済み（評価 #22）**: `evaluator.ts:evaluateTypeIntrinsic` の非リテラル数値分岐に `Number.isInteger(val)` + Integer/Long 範囲チェックを追加。`Dim v As Variant : v = 42 : TypeName(v)` → "Integer"、`VarType(v)` → 2（vbInteger）が正しく返るようになった。`v = 3.14` の浮動小数点は引き続き "Double"。レグレッションテスト: `tests/spec/typename.test.ts` Bug 22-1 ブロック。
77. **`Dir()` はグローバル状態を共有する**: `Dir(pattern)` を呼ぶと内部ポインターがリセットされる。ネストした `Dir(pattern2)` 呼び出しで元の列挙が中断される。実 VBA と同じ仕様準拠の動作。ループ内でサブルーチンを呼ばずに `Dir()` を使うこと。
78. **`DatePart` の第3・第4引数（firstdayofweek / firstweekofyear）は未対応** (Bug #25-6): `DatePart("ww", d, 2)` → Error 450。第2引数まで動作する。`WeekdayName` は第3引数が動作するため非対称。
79. **`LenB` / `AscB` / `ChrB` は未実装** (Bug #25-1〜3): `MidB` は実装済みだが、同じバイト関数グループの他の関数が未実装。`LenB("ABC")` → Error 35。
80. **`Dim b() As Byte : b = str` は修正済み**（評価 #37・Bug #25-4）: 文字列を UTF-16LE のコード単位へ展開する。`"Aあ"` は `41 00 42 30`、下限 0・上限 3 になる。
81. **`Split(str, delim, limit)` の第3引数（limit）は未対応** (Bug #25-5): `Split("A B C", " ", 2)` → Error 450。第2引数まで動作する。
73. **`NPV` に 1-based VBA 配列を渡す場合は修正済み（Bug 24-1）**: `Dim flows(1 To N)` で宣言した配列も `NPV(rate, flows)` で正常動作する。`vbaBase` プロパティを参照して基底インデックスを取得する実装に修正済み。
74. **`Rate` は初期推定値（guessRate）が悪いと別の根に収束する**: `Rate(nper, pmt, pv, 0, 0, guessRate)` の最後の引数 `guessRate` には解に近い値を渡すこと。年率 2% のローンなら `guessRate=0.001`（月率 0.1%）が適切。`guessRate=0.1`（月率 10%）のような遠い値では別の根に収束する。これは VBA 仕様通りの動作（Newton-Raphson 収束の性質）。
75. **`Mid(s, i, n) = val`（代入形式）は実装済み**: 関数形式と同様に正常動作する。`s` の長さは変わらず、指定位置から `len(val)` 文字が上書きされる。
76. **`LSet` / `RSet` は固定長文字列に対して正常動作**: `Dim s As String * 10: LSet s = "ABC"` → `"ABC       "`（左詰め）、`RSet s = "XYZ"` → `"       XYZ"`（右詰め）。
72. **JS 配列を VBA Variant パラメーターに渡すと VarType=8204（vbArray+vbVariant）になる**（評価 #22）: `run('InspectVariant', [[1,2,3]])` → `TypeName="Variant()", VarType=8204, IsArray=True`。`8204 = 8192(vbArray) + 12(vbVariant)`。実 VBA でも配列の VarType はこのビット OR 形式のため、これは正しい動作。
92. **`ReDim Preserve` で UDT 配列を拡張すると新インデックス要素が未初期化（Bug 28-1）**: `ReDim Preserve n(0 To 1)` で添字 1 の UDT 要素が `undefined` のまま残り `n(1).Value = 2` が Error 424 になる。`Long`/`String` の通常配列では同じ操作が正常動作するため非対称。回避策: Preserve を使わず一時配列に手動コピーして置き換えるか、`ReDim n(0 To N)` 後に手動で各要素を `Set` / 初期化する。
93. **`Function` の戻り値と ByRef 引数書き戻しの両方を使う場合の注意**: `Function SafeDivide(a, b, ByRef errMsg)` を `r.run('SafeDivide', [10, 0, ''])` で呼ぶと、戻り値は `run()` の返り値に入り、`errMsg` の書き戻しは `args[0]` ではなく `args[2]`（第3引数）のインデックスに入る。ByRef 書き戻しは引数の元の位置（0-based インデックス）に対応する。`r.run()` が返す配列は `[戻り値, arg0書き戻し, arg1書き戻し, ...]` のように見えるが、実際は呼び出し時に渡した args 配列が直接書き換えられる（`args` オブジェクトへの ByRef 書き戻し）。README の ByRef 例が Sub のみのため Function との組み合わせが分かりにくい。
82. **`Put #` / `Get #` の基本バイナリ I/O は修正済み**（評価 #36・Bug 26-1/2/4/5）: `Byte`/`Integer`/`Long` はリトルエンディアン、文字列は CP932、固定長文字列のみを含む UDT は宣言順の連続バイト列で入出力する。未対応: 配列・可変長文字列を含むUDT、Date、Currency、Single/Double、Random レコード長の実行セマンティクス。
83. **`Open path For Random As #n Len = recLen` はパースエラー**（評価 #26・Bug 26-3）: `Len = <expr>` 節が `parser.ts:parseOpenStatement`（行 977）で未消費。`As #n` の直後にリターンするため `Len` が「ステートメントの後の予期しないトークン」になる。
84. **`GetAttr(path)` は未実装（Error 35）**（評価 #26・Bug 26-7）: `evaluator.ts` に `setattr`（行 999）は stub 登録済みだが `getattr` は未登録。`option-explicit-checker.ts`（行 78）には登録済みのため Option Explicit 違反にはならないが実行時 Error 35。
85. **`FileLen` / `FileDateTime` / `Kill` は正常動作**（評価 #26）: VFS 上のファイルに対してそれぞれ正常に動作する。`Kill path`（括弧なしステートメント形式）も正常。
86. **`Error(n)` 関数の主要コードは正しいメッセージを返す**（評価 #26）: Error(5)="Invalid procedure call or argument" / Error(6)="Overflow" / Error(9)="Subscript out of range" / Error(11)="Division by zero" / Error(13)="Type mismatch" / Error(53)="File not found" / Error(91)="Object variable not set" はすべて正確。未登録番号（Error(0)/Error(7)/Error(14)/Error(999) など）は "Application-defined or object-defined error" にフォールバック。
94. **クラスイベント（`Public Event`/`RaiseEvent`/`Dim WithEvents obj As Class`/`Private Sub obj_EventName()`）は完全動作する**（評価 #29）: イベント引数付き・複数イベント・複数拍のイベント列まで正確。未実装と思われがちだが実装済み。
95. ~~**`eval()` で引数なしクラスメソッドを呼ぶときは `()` を付けること**（評価 #29・Bug 29-F）~~ → **修正済み**（Bug 29-F/G）: `eval('a.Increment')` / `eval('Dim g As New Counter : g.Increment')` どちらも正常に Sub を呼ぶようになった。
96. ~~**手書き `.cls` のヘッダーは「完全」か「なし」のどちらかにすること**（評価 #29・Bug 29-H）~~ → **修正済み**（Bug 29-H）: `VERSION 1.0 CLASS` 行だけ（BEGIN/END なし）でも本体が正しくロードされる。
