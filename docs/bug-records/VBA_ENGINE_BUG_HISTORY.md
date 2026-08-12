---
type: Record
title: VBAエンジンバグ履歴
description: VBA仕様との差異として過去に発見・修正したエンジンバグの履歴。
tags: [vba, bugs, history, specification]
status: stable
---

# VBAエンジンバグ履歴

この文書は、`TODO_SPEC.md` から分離した過去の仕様バグと修正履歴を記録する。
現在の仕様準拠状況と未実装・部分実装・恒久的制限は
[TODO_SPEC.md](../todo/TODO_SPEC.md)を参照する。
個別バグの現在の状態、真因分析、横展開、回帰テストは
`evaluation/findings/BUG-*.md`を正本とする。

以下の記録では、既存のBug ID、修正内容、テスト名を維持している。

## Parser バグ修正

- ✅ **仕様バグ修正（2026-06-25）: `Function Foo() As String()`（配列戻り値型）の `()` が読み捨てられる** | `function-array-return-type.test.ts`
  - 原因: 戻り値型パース（`As <Type>`）が型名トークンのみを読み、後続の `()`（配列マーカー）を見ていなかったため `returnType` が単なる `"String"` になり、`()` トークン自体は本体パース側で空文として黙って消費されていた
  - 症状: `evaluator.ts` の戻り値代入処理が戻り値変数をスカラー String 型と誤認し、`coerceToDeclaredType` 経由で配列に `CStr()` を適用 → `IsArray(GetArr())` が `False` になり、配列要素アクセスで Error 424
  - 修正: `()` を検出したら `ProcedureDeclaration.returnsArray` フラグを立て、`evaluator.ts` 側は `returnsArray` のときスカラー型 coercion を適用しないよう修正（`execProcBody`/`callClassMethod` 両方）

- ✅ **仕様バグ修正（2026-06-26）: クラスフィールドの固定長配列が Empty のまま初期化されず、外部・内部アクセスともに Error 9/438 になる** | `class-module.test.ts`
  - 原因1: `createInstanceFromDef` がフィールドをループ初期化する際に `decl.isArray` チェックがなく、固定長配列が `vbaEmpty` のまま `instanceEnv` に格納されていた。ローカル変数の `evaluateVariableDeclaration` では `decl.isArray → createMultiDimArray()` が正しく呼ばれていた
  - 原因2: 読み取り側の `evaluateCallExpression`（VBA クラス分岐）でも、プロシージャが見つからない場合に配列フィールドへのインデックスアクセスが未実装だった（Error 438）
  - 原因3: 書き込み側の `evaluateAssignmentToVariable`（`obj.Field(i) = val` 形式）でも同様に配列フィールドへの代入が未実装だった
  - 修正: ① `createInstanceFromDef` に `isArray` 分岐を追加（`evaluateVariableDeclaration` と同等の処理）、② 読み取りパスに配列フィールドのインデックス解決を追加、③ 書き込みパスに同様の処理を追加

- ✅ **仕様バグ修正（2026-06-26）: `run()`/`eval()` の戻り値 `Nothing` が JS Symbol のままで、テンプレートリテラルで TypeError / if 判定が誤る** | `test-libs/test-runner.ts`
  - 原因: `vbaNothing = Symbol('vbaNothing')` が `VBARunner.run()`/`eval()` の公開 API 境界で変換されていなかった。Symbol はテンプレートリテラルへの暗黙変換が禁止されており、オブジェクト同様に truthy なため `if (!result)` での Nothing 検出も誤動作した
  - 修正: `VbaBoolean → boolean` 変換と同様に、戻り値が `vbaNothing` のとき `null` に変換するよう `run()`/`eval()` に追加

- ✅ **仕様バグ修正（2026-06-26）: `Dim w As New T` を一度もアクセスせず `Set w = Nothing` すると `Class_Initialize`/`Class_Terminate` が呼ばれない** | `circular-reference-terminate.test.ts`
  - 原因: vba-runner は `Dim w As New T` を遅延インスタンス化（初回アクセス時に生成）で実装していた。`Set w = Nothing` はオブジェクトへの参照とみなされ、実 VBA ではここでインスタンス化が発生してから Nothing 代入になるが、プレースホルダーに対して `triggerTerminate` を呼んでも `__vbaClass__` がないため何もしなかった
  - 修正: `Set x = Nothing` の代入処理（`evaluateSetStatement`）で、`oldVal` が auto-instance プレースホルダーのときは `instantiateClass()` で先にインスタンス化してから `triggerTerminate` を呼ぶよう修正

- ✅ **Fix: `New Collection` — キーワードトークンのクラス名が `New` 式で拒否される** | `new-keyword-classname.test.ts`
  - 原因: `New` 式のパーサーが `classNameToken.type !== Identifier` をチェックしており、`Collection` のように VBA 組み込みキーワードとして字句解析されるクラス名を拒否していた
  - 症状: `New Collection` を含む関数全体がエラーリカバリで消滅し、後続のプロシージャが誤ってトップレベルにリークする
  - 修正: `isNameToken()` ヘルパーを追加し、`Identifier` またはキーワード範囲のトークンを許可。`New Scripting.Dictionary` のようなドット付きクラス名も正しく捕捉するよう拡張

- ✅ **Fix: `As VBA.Collection` — モジュール修飾型の戻り値型が半分しかパースされない** | `module-qualified-type.test.ts`
  - 原因: 関数・Subの戻り値型 `As <Type>` のパーサーがトークンを1つだけ読んでいたため、`As VBA.Collection` は `VBA` だけを returnType とし、`.Collection` が本体ストリームに残った
  - 症状: 残った `.Collection`（キーワード）が本体パースでエラーとなり、関数宣言全体が消滅して `[]` になる
  - 修正: `As` 後にドットが続く場合は `Module.Type` 形式で読み切るよう修正。変数宣言の `Dim x As Module.Type` も同様に対応

- ✅ **Fix: `.Collection` — implicit-With のプロパティ名がキーワードのとき失敗する** | `keyword-property-access.test.ts`
  - 原因: `With obj: .Property = x` のパーサーが `propToken.type !== Identifier` をチェックしており、`.Collection` のようにプロパティ名がキーワードトークンのとき例外を投げていた
  - 症状: `With` ブロック内でキーワード名プロパティを使うと、含む Sub/Function 全体がエラーリカバリで消滅する
  - 修正: `isNameToken()` を使い、`Identifier` またはキーワード範囲のトークンを許可

- ✅ **偽陽性: `MySub()` / `MySub(arg)` を文として書いてもエラーにならない** | `tests/spec/vba_compile_error.test.ts`
  - 実 VBA: "ステートメントの末尾が正しくありません" コンパイルエラー
  - 修正: Pass1 で `CallExpression` を `CallStatement` に変換するブランチで、カッコ付き Sub 呼び出しをエラーとするよう修正

- ✅ **偽陽性: `v = MySub` / `v = MySub()` を式文脈で書いてもエラーにならない** | `tests/spec/vba_compile_error.test.ts`
  - 実 VBA: "Function または変数が必要です" コンパイルエラー
  - 修正: Pass2 (`resolveIdentifiers`) で代入右辺・式文脈で Sub を参照した場合に compile error を発生させるよう修正

## Evaluator バグ修正

- ✅ **Fix: モジュールレベル配列の境界式で `Const` 定数を参照すると境界が `upper=-1` になる** | `dim-const.test.ts`
  - 原因: Pass 1（`evaluateModule`）は `ConstDeclaration` のみスキップしていたが、配列境界付き `VariableDeclaration` は即時評価していた。境界式で `env.get("W")` を呼ぶと `W` が未定義のため `vbaEmpty`（null → 0 に数値変換）が返り、`upper = W - 1 = -1` の空配列が生成された
  - 症状: `Const W = 3: Dim a(0 To W - 1) As Integer` の後で `a(2) = 5` が `Subscript out of range (Error 9)` になる
  - 修正: Pass 1 のアーキテクチャを「シンボルテーブル構築のみ」に統一。配列境界付き Dim を `pendingArrayDecls` に退避、モジュールレベル実行文を `pendingTopLevel` に退避し、`resolveIdentifiers`（Pass 2）で Const 確定後にそれぞれ評価・実行する。`resolveIdentifiersDone` フラグで以降の `evaluateModule`（`evalExpression` 等）は即時実行に切り替える

- ✅ **Fix: `With` ブロック内で JS prototype 上の getter/メソッドが解決できない** | `with-statement.test.ts`
  - 原因: `evaluateImplicitWithObjectExpression` が `Object.keys()`（own enumerable プロパティのみ）と `hasOwnProperty()` で検索していたため、TypeScript class の `get accessor` やメソッドといった prototype チェーン上のメンバーを `.Prop` 構文で発見できなかった
  - 症状: `With obj: .Replacement.Text = "bar"` のように、`.Replacement` が prototype 上の `get accessor` として定義されている JS オブジェクトに対して "Object required (424)" が発生する。`obj.Replacement.Text = "bar"` は `evaluateMemberExpression` 経由で prototype を辿るため正常に動く（非対称）
  - 修正: `evaluateImplicitWithObjectExpression` を `resolveObjectMemberKey`（prototype チェーンを辿る）に統一。`evaluateMemberExpression` と同じロジックを使うことで非対称が解消された

- ✅ **Fix: `Not` 演算子・`vbaToBoolean` が JS ネイティブ `boolean` を変換できない** | `excel-stub.test.ts`, `access-stub.test.ts`
  - 原因: `Not` 演算子は `vbaTrue`/`vbaFalse`（VbaBoolean）と数値のみ処理し、JS ネイティブ `true`/`false` に対して `~true = -2`（truthy）を返していた。`vbaToBoolean` も JS boolean を Type mismatch エラーにしていた
  - 症状: JS モックオブジェクトの boolean ゲッター（例: `MockRecordset.EOF = true`）を VBA で `Do While Not rs.EOF` に使うと無限ループになる。`Dim x As Boolean: x = Application.ScreenUpdating` で Type mismatch が発生する
  - 修正: `Not` 演算子に `typeof argument === 'boolean'` ケースを追加。`vbaToBoolean` も同様に JS boolean を `vbaTrue`/`vbaFalse` に変換するよう対応

- ✅ **Fix: `Dim a(n)` 配列サイズに変数を指定してもエラーにならないバグ** | `dim-const.test.ts`
  - VBA 仕様: `Dim` 文の配列境界式は定数式でなければならない。変数参照は "Compile error: Constant expression required" になる
  - 修正: vb* 組み込み定数を `env.setConstant()` で登録し `isConstant()` で識別可能に。`validateConstantExpr()` を追加し Dim 境界を検証
  - モジュールレベル: `resolveIdentifiers()`（Pass 2 resolve）で検出。プロシージャ内: `checkConstantArrayBoundsInProc()`（`precheckProc`）で AST 走査により検出
  - `ReDim` は変数指定可能・実行時評価のまま変更なし

- ✅ **Refactor: `assertCompileErrorPreproc` を `precheckProc` 専用に構造化** | `vba_compile_error.test.ts`
  - 変更前: `throwAction` が `callProcedure`（precheck + 実行の両方を含む）だったため、実行時エラーでもパターンが合えば通過していた
  - 変更後: `checkProcedure()` を public API として公開し、`throwAction` を `checkProcedure` に変更。`precheckProc` のみを実行し本体は実行しない
  - 併せて以下の AST 静的チェックを `precheckProc` に追加（これらが実行時エラーになっていたため）:
    - `checkDuplicateDimInProc`: 同一スコープ内の重複 Dim 宣言
    - `checkGoToLabelsInProc`: GoTo/GoSub の未定義ラベル（`If False Then GoTo` のような未到達コードも検出）
    - `checkCallArgCountsInProc`: ユーザー定義プロシージャへの引数数不一致

- ✅ **Fix: プロシージャの後に書いた Dim/Const/Type/Enum がコンパイルエラーにならない** | `vba_compile_error.test.ts`（`module_level_dim_after_procedure` 等）, `module-level-toplevel-after-procedure.test.ts`
  - VBA 仕様: モジュールレベル宣言（Dim/Const/Type/Enum 等）はすべてのプロシージャより前に書く必要があり、プロシージャ（End Sub/End Function/End Property）の後に書くと「End Sub、End Function または End Property 以降には、コメントのみが記述できます」というコンパイルエラーになる
  - 修正: `resolveIdentifiers`（Pass 2）に、モジュール内で最初の `ProcedureDeclaration` より後に他の文（宣言文・実行文どちらも）が現れたらエラーにするチェックを追加
  - vba-runner 拡張: `evalVBASingle`/`evalVBAModules` 用のテストスクリプトでは、Dim をプロシージャの後に書いたり、For/Next 等の実行文をモジュールレベルに書く書き方が広く使われている（REPL・簡易テスト用）。これを壊さないよう `Evaluator` に `allowTopLevelStatements`（デフォルト `true`）オプションを追加し、true の間は上記チェックを無効化。`false` を指定すると標準 VBA 相当の挙動になる
  - `CompileError.bas` の自動テスト生成器（`compile-error-generator.ts`）に、ケースごとに `Evaluator` のオプションを上書きできる `EVAL_OPTIONS:` メタフィールドを追加。また `RUNNER: TBD` 調査用の `captureCompileErrorMessage()` を追加（既存の `assertCompileErrorPrerun` 等は pattern `/.+/i` に対して常にマッチしてしまい実際のエラーメッセージを取得できていなかったため）


## サンプルコード検証で発見したエンジンバグ（修正済み / 既知制約）

`sample/workspace/` の Shop サンプル（Product.cls / Inventory.cls / ShopForm.cls / ShopMain.bas）を実行した際に発見。

### ✅ B-1: `parseAsClass` 時にクラス内 `Private Const` が実行時に 0 として解決される（修正済み）

- **修正**: `parser.ts` の `parseClassBody` に `KeywordConst` ケースを追加。`evaluator.ts` の `createInstanceFromDef` でクラスレベルの Const 宣言を `instanceEnv` に評価・登録。
- | `parse-as-class.test.ts` (Test 9: parseAsClass: B-1 — Private Const クラス内参照)

### ✅ B-2: クラス内プライベートメソッドを同クラスメソッドから呼べない（修正済み）

- **修正**: `evaluateCallExpression` で `getProcedure` より先に `Me.__classDef__.procedures` を検索する。VBA の正しい名前解決順序（クラス自身のスコープ → グローバル）を実装。
- | `parse-as-class.test.ts` (Test 10: B-2 — クラス内プライベートメソッド呼び出し, Test 11: B-2 — クラス自身のスコープがグローバルより優先)

### ✅ B-3: `Class_Terminate` の参照カウント制御

- **実装**: クラスインスタンスに `__refCount__` を保持し、`Set` 代入・引数渡し・スコープ終了で参照数を加減算する。参照数が 0 になったときだけ `Class_Terminate` を実行する。
- **保証**: 同じインスタンスを別変数、Dictionary、Collection が保持している間は、いずれかの変数を `Nothing` にしても早期 Terminate しない。
- | `circular-reference-terminate.test.ts`

### ✅ B-4: UDT 固定長配列要素へのフィールド代入が Error 91（修正済み）

- **修正**: `evaluator.ts` に `fillArrayWithUdtInstances` を追加。固定サイズ UDT 配列の各要素を `instantiateType` で個別に初期化する。
- | `udt-type.test.ts` (Test 8: B-4 — 各要素が独立したインスタンス, Test 9: B-4 — 要素間で参照が共有されていない)

### ✅ B-5: クラス内で `Me.Property = value` が Error 91（修正済み）

- **修正**: `parser.ts` の `parseStatement` で識別子ブランチに `TokenType.KeywordMe` を追加。`Me.Prop = val` 形式の代入文が `ImplicitWithObjectExpression`（`.Prop = val`）と誤認識されていたのを修正。
- | `parse-as-class.test.ts` (Test 12: B-5 — Me.Property 代入がクラス内で動作する)

### ✅ B-6: クラス内 `Private WithEvents` フィールド宣言が無視される（修正済み）

- **修正**: `parser.ts` の `parseClassBody` で `KeywordWithEvents` トークンをフィールド宣言開始として認識するよう条件を拡張。`Private WithEvents m_inv As Inventory` のような宣言がフィールドとして登録されず、`Class_Initialize` 内の `Set m_inv = New Inventory` がインスタンス環境に反映されなかった。
- | `parse-as-class.test.ts` (Test 13: B-6 — Private WithEvents フィールドが Class_Initialize で初期化される)

### ✅ B-7: `Event` / `RaiseEvent` でキーワードをイベント名として使えない（修正済み）

- **修正**: `parser.ts` の `parseEventDeclaration` / `parseRaiseEventStatement` で `isNameToken` チェックを追加。`Public Event Open()` / `RaiseEvent Open` のように VBA キーワードをイベント名として使うケース（Excel の Workbook.Open イベント等）が "Expected identifier" エラーになっていた。
- | `raiseevent.test.ts` (既存テストが回帰なしで通過することを確認)

### ✅ B-8: `Function Open/Close()` はモジュールレベルでは仕様上コンパイルエラー（修正済み）

- **経緯**: 過去に `Function Open()` / `Function Close()` をプロシージャ名として受け入れる実装（B-8 fix）を行ったが、Open / Close は `statement-keyword` = `reserved-identifier` （§3.3.5.2）のためモジュールレベルの宣言名には使えない。
- **修正**: `parseProcedureDeclaration` に `STATEMENT_KW_RESERVED` チェックを追加し、これらをコンパイルエラーとして正しく reject するよう変更。クラスモジュールのメンバーとしての `obj.Open` / `obj.Close` はメンバーアクセス（unrestricted-name）なので引き続き有効。
- | `contextual-keyword-as-identifier.test.ts` (Section 6: statement-keyword をプロシージャ名に使うとコンパイルエラー)

### ✅ B-9: クラスの Private/Public フィールドの既定値が型によって Empty のままになる（2026-06-26 修正済み）

- **症状①（UDT）**: `Private mStats As StatBlock`（UDT 型）のようなクラスフィールドが既定値 `Empty` のまま初期化され、`Class_Initialize` 内で `mStats.Strength = 5` のようなメンバー代入をすると `Error 91: Object variable or With block variable not set` になる
- **症状②（Boolean 等）**: `Public IsActive As Boolean` が既定値 `False`(0) ではなく `Empty` になり、`If f.IsActive Then` が意図通り動かない。`Currency`/`Byte`/`LongLong`/`LongPtr` も同様に `0` ではなく `Empty` のままだった
- **症状③（クラス名型フィールド）**: `Public Item As Pub`（クラス型フィールド）が既定値 `Nothing` ではなく `Empty` になり、`Is Nothing` が常に `False` を返す
- **原因**: `Dim` 変数の既定値初期化（`evaluateVariableDeclaration`）は `instantiateType()` での UDT 対応や `boolean`/`currency`/`byte`/`longlong`/`longptr`/クラス型（→Nothing）を含む型集合を持つが、クラスフィールドの初期化（`createInstanceFromDef`）は別の独立したコードパスで、`string`/`integer`/`long`/`double`/`single` の数種類しか default 値を設定していなかった
- **修正**: `createInstanceFromDef` のフィールド既定値判定を `evaluateVariableDeclaration` と同じ型集合に揃え、UDT 型は `instantiateType()` で、クラス名型は `vbaNothing` で初期化するよう変更（`Date`/`Variant` は `Dim` 側でも `Empty` のままが既定の挙動のため対象外）
- | `udt-type.test.ts` (Test 10〜12), `class-module.test.ts`（Boolean/Currency/Byte フィールドの既定値）, `class-object-field-nothing-default.test.ts`

### ✅ B-10: `Set obj.Field = New X` での WithEvents フィールドへの外部代入でイベントハンドラーが配線されない（2026-06-26 修正済み）

- **症状**: `Set s.Source = New Pub`（クラス外部からのメンバーアクセス形式の代入）で WithEvents フィールドに代入しても、イベントハンドラーが一切ワイヤリングされず、イベントが静かに発火しない。`s.Source Is Nothing` は正しく `False`（代入自体は成功している）だが、`RaiseEvent` を含む操作をしても `Source_Ping` ハンドラーが呼ばれない
- **原因**: `evaluateSetStatement` の WithEvents バインディングロジックが `Identifier` LHS ブランチ（`Set Source = New Pub` の形）にしか存在せず、`MemberExpression` LHS ブランチ（`Set s.Source = New Pub` の形）には対応するコードがなかった
- **修正**: バインディングロジックを `bindWithEventsHandlers()` ヘルパーに抽出し、`MemberExpression` LHS ブランチでも `instanceEnv.isWithEvents(propName)` を確認して配線するよう拡張。`classDef`/`instance` にはアサイン先のオブジェクト自身（`obj`）を渡す
- | `raiseevent.test.ts` (Test 5: member access 経由の WithEvents フィールド代入でもハンドラーが配線される)

### ✅ Bug 29-A〜H: 評価 #29 で発見された 8 件のバグ（2026-07-17 修正済み）

- ✅ **Bug 29-A: 型宣言サフィックス付き `Dim n&` / `Dim s$` 等が機能しない**
  - 修正: `parser.ts` の `parseDimStatement` でサフィックス（`%&!#@$^`）を identifier から除去し、型情報を `objectType` に変換。`parsePrimary` でも同じサフィックスを除去して変数参照が一致するよう修正
  - | `type-system.test.ts` (Bug 29-A: 型宣言サフィックス)

- ✅ **Bug 29-B: `^` 演算子が右結合（VBA は左結合）**
  - 修正: `parser.ts:parseExponentiation` の右辺を `parseUnary()`（→ `parseExponentiation` 再入）から単項演算子のみ処理する直接 `parsePrimary()` に変更。`2 ^ 3 ^ 2` = 64 になる
  - | `operators-extra.test.ts` (Bug 29-B: ^ 左結合)

- ✅ **Bug 29-C: 負の底の非整数べき乗が `NaN` を静かに返す（Error 5 にならない）**
  - 修正: `evaluator.ts:case '^'` に `isNaN(result)` チェックを追加し、NaN の場合は Error 5 を throw
  - | `operators-extra.test.ts` (Bug 29-C: (-8)^0.5 → Error 5)

- ✅ **Bug 29-D: `Null & Null` が `""` を返す（Null を返すべき）**
  - 修正: `evaluator.ts:op === '&'` ブランチに「両辺ともに Null なら即 Null を返す」チェックを追加。片側のみ Null の場合は従来通り `""` 扱い
  - | `operators-extra.test.ts` (Bug 29-D: Null & Null → Null)

- ✅ **Bug 29-E: `Array()` が `Option Base 1` を無視する（常に LBound=0 を返す）**
  - 修正: `builtins.ts:array` 登録で `ctx.arrayBase` 分の filler スロットを先頭に追加（`UBound = length - 1` に合わせて `UBound` の計算式も `vbaBase + a.length - 1` → `a.length - 1` へ修正）
  - | `option_base.test.ts` (Bug 29-E: Array() が Option Base 1 を反映する)

- ✅ **Bug 29-F: `eval('a.Increment')` が silent no-op（Sub が実行されない）**
  - 修正: `evaluator.ts:evaluateMemberExpression` の「引数なし Proc 呼び出し」検索条件から `&& p.isFunction` を除去（Sub も呼ぶようにする）
  - | `class-module.test.ts` (Bug 29-F)

- ✅ **Bug 29-G: `eval('Dim g As New Counter : g.Increment : ...')` が Error 450**
  - 修正: `parser.ts` のコール文引数チェックから `TokenType.OperatorColon` を除外（`:` 文区切りを引数区切りと誤認識して `MissingArgument` を生成していた）
  - | `class-module.test.ts` (Bug 29-G)

- ✅ **Bug 29-H: VERSION 行のみで BEGIN/END なしの `.cls` ファイルが本体を全消去する**
  - 修正: `preprocessor.ts:stripVBAFileHeader` で BEGIN 行がある場合のみ BEGIN/END ブロックを除去するよう修正
  - | `preprocessor-cls-header.test.ts` (Bug 29-H)

### ✅ Bug CM〜CT: 監査バッチ（2026-07-17 修正済み）

- ✅ **Bug CM: `Val(42)` が 0 を返す（非文字列引数を String 強制変換してから解析すべき）**
  - 修正: `val` 実装で `typeof s !== 'string'` のとき `vbaToString(s)` で変換してから数値解析
  - | `conversion-int-val.test.ts`

- ✅ **Bug CN: `StrComp("a", "b", Null)` がエラーにならず Binary 比較結果を返す**
  - 修正: Compare=Null のとき `ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)` (error 94)
  - | `strcomp.test.ts`

- ✅ **Bug CO: `Switch(False, "A", True)` が undefined を返す（奇数個引数で Error 5 を投げるべき）**
  - 修正: `args.length % 2 !== 0` のとき `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` (error 5)
  - | `choose_switch.test.ts`

- ✅ **Bug CP: `Round(1.5, Null)` が JS TypeError でクラッシュする（Error 13 を投げるべき）**
  - 修正: `digits === vbaNull` のとき `ctx.throwError(VbaErrorCode.TYPE_MISMATCH, ...)` (error 13)
  - | `builtins.test.ts`

- ✅ **Bug CQ: `Join("notanarray", ",")` が String 化・`Join(Array(1, Null, 3), ",")` が JS TypeError**
  - 修正: 非配列引数は error 13、Null 要素も error 13 に統一
  - | `split-join.test.ts`

- ✅ **Bug CR: `DateAdd("zz", 1, ...)` が無効 Interval でエラーなし（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CS: `DateDiff("zz", d1, d2)` が無効 Interval で 0 を返す（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CT: `DatePart("zz", d)` が無効 Interval で 0 を返す（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CU: `DateSerial(10, 1, 1)` が year 1910 を返す（仕様: 0-29 → 2000-2029）**
  - 修正: `new Date()` 呼び出し前に 2桁年を4桁年に変換（JS legacy behavior 回避のため `setFullYear()` も追加）
  - | `datetime.test.ts`

- ✅ **Bug CV: `Weekday("2025-01-01", Null)` が JS TypeError でクラッシュ（VBA Error に変換すべき）**
  - 修正: `firstdayofweek === vbaNull` のとき `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` (error 5)
  - | `datetime.test.ts`

- ✅ **Bug CX: `Hex(-1)` が "FFFFFFFF"（32ビット）を返す（仕様: "FFFF"、16ビット4文字）**
  - 修正: -32767〜-1 の範囲は `(n & 0xFFFF).toString(16).padStart(4)` で16ビット処理
  - | `conversion-int-val.test.ts`

- ✅ **Bug CY: `Oct(-1)` が "37777777777"（32ビット）を返す（仕様: "177777"、16ビット6文字）**
  - 修正: -32767〜-1 の範囲は `(n & 0xFFFF).toString(8).padStart(6)` で16ビット処理
  - | `conversion-int-val.test.ts`

- ✅ **Bug CZ: `StrReverse(Null)` が vbaNull を返す（仕様: エラー発生）**
  - 修正: Null 入力で `ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)` (error 94)
  - | `strreverse.test.ts`

- ✅ **Bug DA: `IsNumeric(#1/1/2000#)` が True を返す（仕様: False — §6.1.2.7.1.8 は Date を数値型リストに含めない）**
  - 修正: `builtins.ts` の `isnumeric` から `|| val instanceof VbaDate` を削除
  - | `builtins.test.ts`

- ✅ **Bug DB: `String(3, 257)` が "āāā" を返す（仕様: §6.1.2.11.1.38 — 数値 > 255 は `character Mod 256` を使用 → chr(1)*3）**
  - 修正: `String.fromCharCode(char)` → `String.fromCharCode(Math.trunc(char) % 256)`
  - | `builtin-strings.test.ts`

- ✅ **Bug DC: `String(3, "")` がエラーを発生しない（仕様: §6.1.2.11.1.38 — 空文字 Character は Error 5）**
  - 修正: 空文字列チェック後 `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)`
  - | `builtin-strings.test.ts`

- ✅ **Bug DD: `LBound(arr, Null)` / `UBound(arr, Null)` が JS TypeError でクラッシュ（VBA Error に変換すべき）**
  - 修正: `if (dim === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `builtins.test.ts`

- ✅ **Bug DE: `LBound(arr, 0)` / `UBound(arr, 0)` が Error 9 を発生しない（次元は1ベース、0は無効）**
  - 修正: `if (dimIndex > 0)` → `if (dimIndex < 0 || dimIndex > 0)` で負値も捕捉
  - | `builtins.test.ts`

- ✅ **Bug DG: `StrConv("abc", Null)` が JS TypeError でクラッシュ（Conversion は VbStrConv 型、Null → Error 94）**
  - 修正: `if (conv === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `strconv_null.test.ts`

- ✅ **Bug DH: `IsObject(Date/Boolean/ErrorValue)` が True を返す（値型はオブジェクト参照ではない）**
  - 修正: `val instanceof VbaDate || VbaBoolean || VbaDecimal || VbaCurrency || VbaErrorValue` は False を返すよう instanceof チェックを追加
  - | `builtins.test.ts`

- ✅ **Bug DI: `Error(Null)` が JS TypeError でクラッシュ（Let-coerce to Long で Null → Error 94）**
  - 修正: `if (n === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `builtins.test.ts`

- ✅ **Bug DJ: `QBColor(Null)` が JS TypeError でクラッシュ（`Number(Symbol)` — Null → Error 94）**
  - 修正: `if (c === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `builtins.test.ts`

- ✅ **Bug DK: `Rnd(Null)` が JS TypeError でクラッシュ（Symbol との比較 `val < 0` が TypeError）**
  - 修正: Null チェックを先頭に追加。`val < 0` を `typeof val === 'number' && val < 0` に変更
  - | `builtins.test.ts`

- ✅ **Bug DL: `Randomize(Null)` が JS TypeError でクラッシュ（`Number(Symbol)` — vbaNull ≠ null）**
  - 修正: `if (val === vbaNull) ctx.throwError(...)` を先頭に追加
  - | `builtins.test.ts`

- ✅ **Bug DM: `RGB(Null, 0, 0)` が JS TypeError でクラッシュ（`Number(Symbol)` — 各引数の Null チェック欠如）**
  - 修正: 全引数の Null チェックを追加
  - | `builtins.test.ts`

- ✅ **Bug DN: `Format(123, Null)` が Format パターンとして "Symbol(vbaNull)" を使い、ゴミ文字列を返す**
  - 修正: pattern === vbaNull のとき Error 94 を throw（Format As String パラメーターは Null 不可）
  - | `builtins.test.ts`

- ✅ **Bug DO: `FormatDateTime(date, Null)` が `Number(Symbol)` JS TypeError でクラッシュ**
  - 修正: namedFmt === vbaNull のとき Error 94 を throw
  - | `builtins.test.ts`

- ✅ **Bug DP: `LeftB/RightB` が負の Length でエラーを出さない、`MidB` が Start=0 でエラーを出さない**
  - 修正: LeftB/RightB に `byteLen < 0 → Error 5` チェック追加、MidB に `byteStart < 1 → Error 5` チェック追加
  - | `builtins.test.ts`

- ✅ **Bug DQ: 財務関数（FV/PV/PMT/NPER/RATE/SLN/SYD/DDB/IRR/MIRR/NPV/IPMT/PPMT）で Null 引数が JS TypeError クラッシュ**
  - `registerFinancialFunctions` に `toNum` ヘルパーを追加し、全財務関数の `Number(arg)` 呼び出しを `toNum(arg)` に統一
  - toNum: Null チェック → Error 94、その後 `Number()` に委譲
  - | `builtins.test.ts`

- ✅ **Bug DR: `CallByName(obj, Null, callType)` / `CallByName(obj, name, Null)` が Error 438 を返す（仕様: Error 94）**
  - `ProcName`/`CallType` 引数の型を `any` に変更し、Null チェックで Error 94 を throw
  - | `builtins.test.ts`

- ✅ **Bug DS: `DeleteSetting` が非存在の App/Section/Key を削除してもエラーを出さない（仕様: 実行時エラー）**
  - App 未存在・Section 未存在・Key 未存在それぞれで Error 5 を throw
  - | `builtins.test.ts`

- ✅ **Bug DT: 組み込み関数の異常値ファジング（`scripts/fuzz-builtins.ts`）で検出した 359 件のクラッシュ**
  - 184 関数 × 12 種の敵対値（Null/Empty/Nothing/""/巨大数/配列 等）= 4361 呼び出しを全数検査し、JS 生例外（TypeError/RangeError/SyntaxError/generic Error）で落ちるケースを 0 件にした
  - 中央修正: `coerce.ts` の `vbaToNumber`/`vbaToString`/`vbaToBoolean` に Nothing → Error 91、配列・非数値 → Error 13、undefined → 0（Empty 相当）を追加。`parseVbaDate` に Nothing → Error 91。`VbaCurrency.fromNumber` に非有限・明確な範囲外 → Error 6（境界付近は従来どおり正確な文字列→BigInt 経路で判定し精度を保持）
  - `Space`/`String` の Long 超過 → Error 6、JS 文字列上限超過 → Error 14（`OUT_OF_STRING_SPACE` を新設）
  - `FormatNumber` 系の桁数 > 255 → Error 5（101〜255 桁は 0 埋めで許容）
  - `FileLen`/`FileDateTime`/`FileCopy` の不存在パス → Error 53（従来は generic Error）
  - builtins.ts 全体の生 `Number()`/`String()` 引数強制を中央ヘルパー（`ctx.toVbaNumber`/`vbaToString`）に置換。evaluator.ts のファイル I/O 系（EOF/LOF/Loc/Seek/Kill/MkDir 等）・COM 系（CreateObject/GetObject）・Interaction 系（MsgBox/Shell 等）も同様
  - 組み込み呼び出し経路に `invokeBuiltin` ラッパーを追加し、素の VbaError オブジェクトを行番号付き `Run-time error 'NN'` 形式に包み直すよう統一
  - | `fuzz-regressions.test.ts`

- ✅ **Bug DU: On Error 組み合わせストレス（`scripts/stress-on-error.ts`）で検出した 2 バグ**
  - ハンドラーモード × エラー位置 × エラー種別 × 構造の 128 シナリオを実行トレースで検証
  - **DU-1**: ネストブロック（For 等）内のエラーを GoTo ハンドラーで処理して `Resume Next` すると、失敗文の次ではなくブロック全体の次へ飛びループが中断された。修正: `executeStatements` がループフレームを保ったままプロシージャ本体のハンドラーを実行し、Resume で同フレームへ復帰（`Resume` は再実行、`Resume Next` は次文、`Resume <label>` は GoTo として脱出、fall-through は正常終了）。単一行 If の節は `inlineClause` フラグで除外し、If 文全体の粒度で Resume する（実 VBA の行粒度に一致）
  - **DU-2**: `Err.Raise` が Description 未指定時に既定メッセージ（§6.1.3.2: Number 対応の Error 関数文字列、なければ "Application-defined or object-defined error"）を設定せず、素の VbaError オブジェクトが枠組みなしで escape していた。`raise()` に既定 Description を追加し、メソッド呼び出し経路と `Error` 文を `invokeBuiltin` で包んで統一
  - 仕様確認: `End Sub` は Err を自動クリアしない（§6.1.3.2 の自動 Clear は Resume / Exit Sub・Function・Property / On Error 文のみ）。エンジンの既存挙動が正しいことをストレスのモデル側で裁定
  - | `on-error-nested-resume.test.ts`

- ✅ **Bug 31-A: `Public/Private Static Sub|Function` がパースエラー（評価 #31）**
  - `parser.ts` の `KeywordPublic/Private` 分岐に `Static` + `Sub|Function|Property` の先読みを追加し、`parseProcedureDeclaration(scope, true)` へ委譲
  - | `static.test.ts`

- ✅ **Bug 31-B: `Erl` が未実装で常に 0 を返す（評価 #31）**
  - 数値行ラベル（`20 x = 1`）の通過時に `lastLineNumberLabel` を記録し、エラー捕捉時に `Err` 状態へ確定。`Erl` 組み込みを登録（引数なし・自動呼び出し可能）
  - `Err.Clear` / `Resume` / `On Error` 文 / `Exit` 系で Err とともにリセット（MS-VBAL に Erl の記載はないため実 VBA の通説に準拠）
  - | `erl-function.test.ts`

- ✅ **Bug 32-A〜32-E: 評価 #32（テンプレートエンジンドメイン）で発見した 5 バグ**
  - 32-A: `Tab(n)` が n+1 桁目に出力するオフバイワン（Print # / Debug.Print 両経路）→ `n-1` 桁までのパディングに修正
  - 32-B: レクサーが `Write #1, #2024/03/15#` の `#1, #` を日付リテラルと誤認 → 日付判定に区切り文字（`/` `-` `:`）の存在を必須化
  - 32-C: `Input #` が引用符内のカンマで分割 → 引用符状態を追跡するフィールド分割に変更
  - 32-D: `Write #` の日付が `#yyyy/mm/dd#` + LF → universal format `#yyyy-mm-dd#` + CRLF に修正（`Input #` に日付復元も追加）
  - 32-E: `LSet` の UDT 間コピーが Error 424 → 同一レイアウトなら位置ベースのフィールドコピー、不一致は明示的な Error 5
  - | `write-input-print-zones.test.ts`, `lset-rset.test.ts`

- ✅ **Bug 33-A〜33-C: 評価 #33（カードゲームドメイン）で発見した 3 バグ**
  - 33-A: クラス内の括弧なし自メンバー参照（Property Get / Function）が silent Empty → Identifier 評価に暗黙 Me.<name> フォールバックを追加
  - 33-B: `Collection.Add` の名前付き `Before:=`/`After:=` が誤バインド → `__vbaParamSpec__` 付与。未知・解決不能な名前付き引数は Error 448（Named argument not found）に
  - 33-C: 未使用 `As New` 変数の ByVal 渡しで実体が呼び出し元に反映されない → 引数評価時に `resolveAutoInstance` で caller 側に実体化
  - | `class-module.test.ts`, `collection.test.ts`, `auto-instance-args.test.ts`

- ✅ **Bug: 括弧付き数値リテラルが Variant サブタイプを失う（実 VBA 差分コーパスの整備で発見）**
  - `v = (2)` / `v = (2) + (2)` の TypeName が Double になっていた（実 VBA は Integer）
  - `resolveNumericSubtype` が `ParenthesizedExpression` を unwrap していなかった
  - | `typename.test.ts`

- ✅ **実 VBA 差分テスト第 1 回実施（scripts/gen-diff-corpus.ts 2337 式を実 Excel VBA と突き合わせ）**
  - 初回 diff: 1608/2337 一致 → 修正を重ねて最終 2291/2337 一致（残り 46 件は環境依存として `scripts/diff-allowlist.txt` へ登録）
  - **型伝播の是正**: 関数呼び出し引数の Variant サブタイプが呼び出し元 env で解決されず Double に落ちていた（`callProcedure` にモジュール修飾・非修飾の両呼び出し経路で `argSubtypes` を伝播）。`Sgn`/`StrComp`/`DatePart`/`Year`/`Month`/`Day`/`Hour`/`Minute`/`Second`/`Weekday` の組み込み戻り型を `BUILTIN_RETURN_TYPES` に追加
  - **Integer/Long 演算のオーバーフロー**: `(32767)+(2)` のような Integer 同士の算術は Long へ暗黙昇格せず Error 6 になる仕様を実装（`evaluateBinaryExpression` にラッパーを追加）。片方の型が静的に不明な場合は誤検出防止のため判定をスキップ（`Sgn(x)*Abs(x)` のようなケースで一度リグレッションし修正）
  - **0 除算**: `0/0` は Error 6（Overflow）、非ゼロ/0 は Error 11（Division by zero）と VBA は使い分ける
  - **Null 伝播の非対称性**: `"abc" - Null` は Error 13（数値変換失敗が先行）だが `"abc" + Null` は Null のまま伝播するという `+` 演算子だけの特例
  - **比較演算子の型変換**: 文字列×数値/Date は数値変換して比較（変換不能は Error 13）。文字列×Boolean は `CBool` 変換して比較（`"7" = True` は True）
  - **Date 演算**: `数値 ± Date` は Date 型維持（`-` は左右どちらが Date でも）、`Date * 数値` は Double、`Date - Date` は Double。範囲外は Error 6
  - **CDate/CCur/CBool の数値文字列解釈統一**: `&H`/`&O`/指数/カンマ区切りの文字列を数値として解釈（中央の `toVbaNumber` 経由に統一）。`CDate(Empty)` はシリアル値 0
  - **`fromVbaDate` の負シリアル値の丸め方向修正**: 整数部（0 方向切り捨て）が日数、絶対値の小数部が時刻という VBA の分解規則に合わせた（`CDate(-2.5)` 等）
  - **`DateSerial` の月/日オーバーフロー**: `date.setFullYear()` で年を強制上書きしていたため月の繰り上げ・繰り下げが year に反映されていなかった。0-99 年の特殊レンジのみ基準年オフセット方式、それ以外は JS の Date コンストラクターにそのまま繰り上げを委ねる
  - **`Format()` の丸めは通常の四捨五入**（銀行家丸めではない）と判明。過去の Bug BG（銀行家丸めへの修正）は誤りだったため訂正
  - **`CStr(Null)` は Error 94**（空文字列ではない）と判明。過去の Bug BK の理解を訂正
  - **`Str()` の VBA 書式**: 15 有効桁・先頭ゼロ省略・E+NN 指数形式を実装
  - **Trim/LTrim/RTrim はスペース（Chr 32）のみ除去**（タブ等は残る）
  - 46 件の残存差分は許容リストへ（コードページ依存の ANSI 文字変換、1937 年以前のタイムゾーン LMT 依存、Single 型の表示精度、`\`/`Mod` での Boolean 型保持等）
  - | `real-vba-diff-regressions.test.ts`

- ✅ **実 VBA 差分テストの未解明差分を深掘り（\`/\`Mod\` の Boolean 保持・CDate の "H.N" 時刻解釈）**
  - 一致 2291/2337 → **2302/2337** に向上（許容リスト 46→35 件）
  - **`\`/`Mod` の Boolean 型保持は「左が Boolean かつ右が文字列」のときだけの非対称規則**と判明: `True \ "7"` は `"7"` が `CBool` 変換され `True(-1)\True(-1)=1`→非ゼロで Boolean True。しかし `"7" \ True`（左右逆）は通常の数値変換で Long -7、`True \ True`（両方 literal で文字列変換なし）は通常の整数昇格で Integer 1。片方向・文字列変換が絡む場合のみ発動する規則
  - **CDate の "H.N" 形式の文字列は「H時N分」の時刻として解釈される**（シリアル値の小数としてではない）と判明: `CDate("2.5")` = 02:05:00。共通ヘルパー `tryParseTimeFractionString`（`vba-types.ts`）として実装し、CDate と `+` 演算子（Date とのペア限定）の両方から利用
  - この時刻解釈規則は **`+` 演算子だけの特殊ルール**: `#date# + "3.5"` は 03:05:00 を加算するが、`-`/`*`/`/`/`\`/`Mod` では同じ `"3.5"` を通常どおりシリアル値 3.5 として解釈する（非対称）。実装時に一度全演算子に広げてリグレッションし、`+` 限定に絞って解決
  - `"M,Y"` 形式（`CDate("1,234")` → 西暦234年1月1日）は差分コーパス第2回で現代年を含む境界を実Excel再検証済み。極端に古い年のタイムゾーン LMT 依存差分のみ許容リストに残置（詳細: `EVAL_LOG.md` #122）
  - | `real-vba-diff-regressions.test.ts`

- ✅ **実 VBA 差分テスト第2回（境界値の追加検証、149 式）**
  - 上記で発見した2規則の境界を狙ったコーパスを追加生成（`\`/`Mod` の Boolean 保持規則で右辺が数値/Empty/Null/Date のケース・他の算術演算子での再現有無、`CDate("H.N")` の無効値/ゼロパディング/区切りなし、`CDate("M,Y")` を現代の年で再検証、`+` の Date 型時刻解釈の周辺ケース）
  - 計 2486/2486 式中 **148/149 件の追加分が初回から一致**（実装した規則の正確さが裏付けられた）。唯一の不一致 `MYDATE_0005`（`CDate("1,234")` を現代年版コーパスにも含めた参考ケース）は既知のタイムゾーン LMT 問題と同一原因で許容登録
  - 最終結果: **2450/2486 一致**（許容リスト 36 件）
  - | `real-vba-diff-regressions.test.ts`

- ✅ **実 VBA 差分テスト第3回（論理演算子・StrConv 拡張、170 式追加、計 2656 式）**
  - `And`/`Or`/`Xor`/`Eqv`/`Imp`（Boolean × 文字列/数値/Null/Empty の型強制、160件）と `StrConv` の日本語変換（ひらがな⇔カタカナ・全角⇔半角の往復・複合フラグ、10件）を追加
  - **Bug 発見・修正1: 論理演算子で片方が Boolean のとき、結果の数値サブタイプが Double にフォールバックしていた**（値自体は正しく、`TypeName`/`VarType` だけがズレていた）。`resolveBinaryExprNumericSubtype` に `and`/`or`/`xor`/`eqv`/`imp` の分岐がなく、`\`/`Mod` と同じ「非整数型は Long に丸めてから演算」規則が適用されていなかった。`toIntType` ヘルパーを `\`/`Mod` と論理演算子で共通化して解決（文字列が絡む場合は Long、数値リテラル/Empty のみなら Integer、実測どおり）
  - **Bug 発見・修正2: `StrConv` の `vbWide`(4)/`vbNarrow`(8) が半角カナ⇔全角カタカナの変換に未対応**（英数字・記号のみ対応で、カナは素通りしていた）。半角→全角は `String.prototype.normalize('NFKC')` で濁点・半濁点含め正確に変換できることを確認して採用。全角→半角は Unicode 正規化では分解できないため、明示的なマッピングテーブル（`FULL_TO_HALF_KANA`）を実装
  - 一致 2450/2486 → **2620/2656**（追加分 170 式中 167 件が初回一致、3件が上記バグで不一致 → 修正後 170/170 一致）。許容リスト件数は変わらず 36 件
  - | `real-vba-diff-regressions.test.ts`

## 追加の仕様バグ修正履歴

- ✅ **仕様バグ修正（2026-06-25）: `Open ... For Append` が追記ではなく上書きになる（サイレントなデータ消失）**
  - 原因: `MemoryFileSystem.openSync()` がファイルハンドルの書き込み開始位置 `pos` を `flags` に関係なく常に `0` で初期化していたため、`writeSync` が `Append` モードでも常にファイル先頭から上書きしていた。2回目以降の `Open ... For Append` のたびに、それまでの追記内容が消える
  - 症状: 例外は発生せず、ログファイル追記のような典型的な用途でデータが静かに失われる。`NodeFileSystem`（実ファイル版、`vba-run` CLI 等で使用）は Node 標準の `fs.openSync` に委譲しているため影響を受けず、`VBARunner` が使う `MemoryFileSystem` のみの問題
  - 修正: `flags === 'a'` の場合、既存ファイルが存在すればその長さを `pos` の初期値にするよう変更 | `filesystem.test.ts`
- ✅ **仕様バグ修正（2026-06-26）: 組み込み実行時エラーの `Err.Description` が `"Run-time error 'N': ... (line X)"` の枠組みテキストを含む**
  - 原因: `throwVbaError()` が TypeScript 側に投げる例外の `.message` 用に枠組み付き文字列を生成しており、`executeStatements` の catch ブロックがその `e.message` をそのまま `errObj.description` に転記していた。`Err.Raise`（`ErrObject.raise()`）経由のカスタムエラーは元々生のメッセージを直接 `description` に入れていたため問題なかったが、組み込みエラー（Type Mismatch・Division by Zero 等）だけ枠組み付きになっていた
  - 修正: `throwVbaError()` が投げる例外に `vbaBareMessage`（枠組みなしの生のメッセージ）を追加で持たせ、`errObj.description` への転記時はそちらを優先する。TypeScript 側に渡る `.message` 自体（`eval()`/`run()` の catch で見えるエラーメッセージ）は枠組み付きのまま変更していない | `err-object-full.test.ts`
- ✅ **仕様バグ修正（2026-06-26）: `vbObjectError` 定数が未実装**
  - 症状: カスタムエラー番号を生成する MS 推奨の慣用句 `Err.Raise vbObjectError + n` が `Option Explicit` 配下で未宣言変数エラーになる
  - 修正: `vbobjecterror = -2147221504`（§6.1.2.10）を組み込み定数として追加 | `err-object-full.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `On Error Resume Next:` / `On Error GoTo Label:` の直後に `:` で後続文を連結すると消える**
  - 原因: `parser.ts` の `parseOnErrorStatement()` が "Resume Next" 部分を読み取るループで `Newline`/`EOF` のみを終端条件にしており、`:`（`OperatorColon`）で止まらなかったため、同一行の `:` 以降の文すべてが label 文字列として読み込まれ AST から消えていた
  - 修正: 既存の `isAtTerminator()`（`Newline`/`EOF`/`OperatorColon` を終端とする）を使うよう変更 | `on-error-colon.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.eval()` 内で `On Error Resume Next`/`On Error GoTo` が機能しない**
  - 原因: `evalExpression()` の複数文フォールバックが `evaluateModule()` の単純な逐次実行ループを使っており、`On Error`（`Resume Next`/`GoTo`/`Resume`）の分岐ロジックを持つ `executeStatements()`（手続き本体実行や `run()` 経由では使われる）を経由していなかった
  - 修正: `evalExpression()` のフォールバックを `executeStatements(program.body, 0)` に変更。さらに `eval()` 呼び出しごとに `errorHandlerLabel`/`errorHandlingMode`/`isInErrorHandler`/`lastErrorIndex`（および `currentSourceModule`）を退避・復元し、`eval()` を独立したトップレベル呼び出しフレームとして扱うよう統一 | `eval-on-error.test.ts`, `eval-after-load.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: クラスメソッド呼び出しが On Error 状態を呼び出し元に漏らす**
  - 原因: `callClassMethod()` が `errorHandlerLabel`/`errorHandlingMode`/`isInErrorHandler`/`lastErrorIndex` を退避・復元していなかった（通常の手続き呼び出し `execProcBody` は退避・復元している）。クラスメソッド内で `On Error` ハンドラーに入った状態（`isInErrorHandler=true`）のまま例外が呼び出し元に伝播すると、呼び出し元の `executeStatements` がこれを引き継いでしまい、呼び出し元自身の `On Error GoTo`/`Resume Next` への分岐が機能せずエラーがそのまま素通りしていた
  - 修正: `execProcBody` と同様に4フィールドを退避・リセット・復元するよう `callClassMethod` を修正 | `class-method-on-error-isolation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.eval()` で定義したプロシージャが `Option Explicit` の静的検査対象にならない**
  - 原因: `resolveIdentifiers()`（Pass 2）は `VBARunner` 構築時にロードしたモジュール群のみを `checkOptionExplicit` の対象にしており、構築後に `eval()` で定義したプロシージャは一度も解析されず `optionExplicitViolations` に登録されなかった
  - 修正: `evalExpression()` の複数文フォールバックで都度 `checkOptionExplicit` を実行し、違反があれば `optionExplicitViolations` に追加登録するよう修正（既存の登録は維持） | `eval-option-explicit.test.ts`
- ✅ **Fix: モジュール修飾呼び出しが呼び出し先のランタイムエラーを握りつぶす** | `runtime-error-trace.test.ts`
  - 原因: `evaluateCallExpression` の `Module.Proc` 検出が `try { callProcedure } catch { フォールスルー }` で**あらゆる例外**を飲み込み、member access に落ちていた
  - 症状: `ModB.DoWork` の本体がゼロ除算（Error 11）を投げても握りつぶされ、未定義オブジェクト `ModB` への member access として Error 91 にすり替わる
  - 修正: `getProcedureFromModule` で**事前に存在確認**し、見つかった場合のみ `callProcedure` を呼ぶ（try/catch を廃止）。呼び出し先のエラーは正しく伝播する
- ✅ **Fix: `VBA.Func()` が標準ライブラリ関数を呼ぶ強制力を持たない** | `vba-stdlib-qualifier.test.ts`
  - VBA 仕様: `VBA.InStr(...)` のような型ライブラリ修飾呼び出しは、ユーザーが同名の関数を定義していても必ず標準ライブラリを呼ぶ
  - 原因: `evaluateCallExpression` の `VBA.X` 処理が callee を非修飾 `X` に差し替えて再帰していたため、`getProcedure`（ユーザー定義優先）の経路に乗っていた
  - 修正: `env.getConst(name)` で `variables`（組み込み関数テーブル）のみを検索し、`procedures`（ユーザー定義）をスキップして直接呼び出す
- ✅ **Fix: `VarType(VBA)` / `VarType(Module1)` のようにプロジェクト名・モジュール名を値として使った場合にエラーにならない** | `namespace-as-value-error.test.ts`
  - VBA 仕様: `VarType(VBA)` → コンパイルエラー「プロジェクトではなく、変数またはプロシージャを指定してください」。モジュール名も同様。
  - 修正: `vba-types.ts` に `VbaNamespaceRef` センチネルクラスを追加し、`VBA` を `kind='project'` で、各モジュール名を `kind='module'` で env に事前登録。`Identifier` 評価時に `VbaNamespaceRef` を検出してエラーを投げる。修飾形式（`VBA.X`、`Module1.Proc`）は早期リターンで正常動作を維持。
  - **MS-VBAL 上の「プロジェクト」の定義・実装詳細**: `docs/implementation/REFERENCE.md` §「外部ライブラリの参照設定と `New ProjectName.ClassName`」を参照。`registerExternalObject('X.Y', factory)` でプロジェクト名 X が `VbaNamespaceRef kind='project'` として自動登録される仕様を含む。
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.run(name, args)` で ByRef パラメーターに代入しても呼び出し元の `args` 配列に反映されない**
  - 「あるべき姿」: VBA の既定の引数渡しは ByRef（明示的な `ByVal` がない限り）。`Sub Increment(ByRef n As Long): n = n + 1` を `run('Increment', [5])` で呼んだ後、呼び出し元が渡した `args` 配列は `[6]` になっているべき（「ステータスコード + メッセージを ByRef の out パラメーターで返す」という VBA で一般的なパターンを TypeScript 側で受け取れる必要がある）
  - 「修正前の現状」: `callProcedure()` は引数を `localEnv` に値としてコピーするだけで、JS の数値・文字列・Boolean はプリミティブで参照を共有しないため、Sub 内で ByRef パラメーターに代入してもエラーは出ないが `args` 配列には一切反映されなかった（`evaluateCallExpression` 経由の VBA-to-VBA 呼び出しは `originalExpr` を使った別の書き戻し機構を既に持っており問題なし。`callProcedure()` 経由＝主に `VBARunner.run()` からの直接呼び出しのみが対象）
  - 修正: `callProcedure()` で `execProcBody()` 呼び出し後、`ParamArray` でも `ByVal` でもないパラメーターについて `localEnv` から最終値を読み、呼び出し元の `args[i]` に書き戻すよう変更 | `run-byref-writeback.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `Join()` が下限付き固定配列（`Dim a(1 To n)`）の先頭に余分な空要素を混入させる**
  - 原因: `Dim a(1 To n)` の物理ストレージは `LBound` 分の隠し添字（添字0など）を含む JS 配列で確保され、`UBound`/`LBound`/添字境界チェックは `vbaBase`/`__vbaDimensions__` を見て隠しているが、`join` ビルトインは `arr.join(del)` を物理配列にそのまま適用していた
  - 修正: `vbaBase`（既定 0）でスライスしてから `join` するよう修正 | `split-join.test.ts`
- ✅ **仕様バグ修正（2026-06-26）: `eval()` の高速パスが式実行中の本物のランタイムエラーを握りつぶし、無関係な構文エラーに化けさせる**
  - 原因: `evalExpression()` は「単一式としてパースできるか」を try/catch で試すが、その catch ブロックが `parseExpressionPublic()` の呼び出しだけでなく後続の実行（`callProcedure`/`evaluateExpression`）まで囲んでいた。式の中で呼んだ関数が `Error 91` 等の本物のランタイムエラーを投げても「パース失敗」として握りつぶされ、文として再解析した際に得られる無関係な構文エラー（または別の挙動）に置き換わっていた
  - 症状: `eval('Foo()')`（`Foo` が内部で実行時エラーを投げる）が `run('Foo', [])` と異なる、誤解を招くエラー（または `Foo()` を裸の statement として再解析したことによる "syntax error"）になる。デバッグ時に本当の原因が一切見えなくなる
  - 修正: パース段階（`parseExpressionPublic()` 呼び出しと `fullyConsumed`/`isStatementAmbiguous` 判定）のみを try/catch で囲み、実行（`callProcedure`/`evaluateExpression`）はその外側で行うよう構造を変更 | `eval-real-error-propagation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval()` に複数行コードを渡すと1行目だけが式として誤評価され、2行目以降が黒く無視される**
  - 原因: `evalExpression()` の「単一式として全体を消費したか」判定が `parseExpressionPublic()` 直後のトークンが `Newline` であることだけを見ており、その先に何があるか確認していなかった。VBA の `=` は代入文とも等価比較式とも解釈できるため、`"x = 10\nDebug.Print 1"` のような入力では1行目 `x = 10` が単独で「等価比較式」としてパースしきれてしまい、直後が改行であることから「全体を消費した」と誤判定していた
  - 症状: `x = 10` が代入ではなく比較として評価され（`x` は未代入のまま）、2行目以降（`Debug.Print 1` 等）は一切実行されない。`eval()` に Sub/Function 定義を含まない複数行スニペットを渡す、ごく自然な使い方で発生する
  - 修正: 直後が `Newline` の場合は、それ以降（複数の `Newline` を挟んでも）が `EOF` であることまで確認するよう変更。単一式＋末尾改行のみという既存の正しいケースは維持される | `eval-multiline-truncation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval()` に裸の代入文（`x = 10` 単体、`arr(1) = "a"` 単体）を渡すと代入が一切実行されない**
  - 原因: 上記と同根。`parseExpressionPublic()` が `x = 10` を「`x` と `10` を比較する `BinaryExpression`」として入力全体ぴったり消費する形でパースしてしまい、高速パスの「単一式として全体を消費した」判定を満たしてしまっていた。代入は一度も実行されず、比較結果の真偽値だけが返っていた
  - 症状: `eval('x = 10')` の直後に `eval('x')` しても未代入のまま。配列要素も同様（`eval('arr(1) = "a"')` の代入が反映されない）。`Dim` と代入を別々の `eval()` 呼び出しに分けるという自然な使い方で発生する
  - 修正: 入力全体を消費した式がトップレベルの `=` を持つ `BinaryExpression` の場合は高速パスを使わず、文として解析・実行するフォールバックに委ねるよう変更（実 VBA でも裸の `lhs = rhs` 文は常に代入と解釈されるため、副作用なしの比較式として読まれる余地はない） | `eval-bare-assignment.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval('x + 1')`（単独）と `eval('... : x + 1')`（複数文の最後）の結果が一致しない**
  - 前提（実 VBA 仕様の確認）: VBA の statement 文法では、識別子で始まり `+`/`-`（および `=`）が続く裸の文は、単独行でも複数文中でも**常に同一の解釈**になる。一方 `<`/`&`/`And` 等の比較・連結・論理演算子は statement としては成立せず Parse error になるため、この曖昧性自体が存在しない（`x < 10` を文として解析するとパースエラー）。検証結果:
    | 入力（文として解析） | 解析結果 |
    |---|---|
    | `x - 1` | `CallStatement`（`x` を引数 `-1` で呼び出す） |
    | `x + 1 + 2` / `x + 1 - 2` | `CallStatement`（深い `+`/`-` チェーンも同様に1つの呼び出しに集約） |
    | `Foo(1) + 1` | `CallStatement`（明示括弧付き呼び出しが左辺の場合。後述の追加修正で式評価に変更） |
    | `x.Bar + 1` | `CallStatement`（メンバアクセスが左辺でも同様） |
    | `1 + x`（識別子が左辺でない） | `LabelStatement` + 別文（数値リテラル開始は行ラベルと解釈され、別カテゴリの曖昧性のため対象外） |
    | `(x) + 1`（括弧で囲む） | Parse error（括弧で明示的に式扱いになり曖昧性が解消される） |
    | `x < 10` / `x & "a"` / `x And y` | Parse error（statement として成立しないため曖昧性なし。式としての評価が唯一の解釈） |
  - 「あるべき姿」と「修正前の現状」:
    | x の種類 | 文脈 | あるべき結果（実VBA仕様） | 修正前の現状 |
    |---|---|---|---|
    | 変数（例 `x=5`） | 単独 `eval('x + 1')` | `x` は呼び出し不可能なため `Error 424` | ❌ `6`（算術加算として誤評価） |
    | 変数 | 複数文中 `... : x + 1` | `Error 424` | ✅ `Error 424`（既に正しい） |
    | Function（必須引数1個） | 単独 | 引数 `1` で呼ばれ、戻り値は文として捨てられる（エラーなし） | ❌ `1`（`x` が暗黙的に `Empty`(0) 扱いされ `0+1` を返す。呼び出しは発生しない） |
    | Function（必須引数1個） | 複数文中 | 同上 | ✅ 同上（既に正しい） |
    | Sub（必須引数1個） | 単独 | 引数 `1` で呼ばれる（エラーなし） | ❌ `1`（同上、呼び出しは発生しない） |
    | Sub（必須引数1個） | 複数文中 | 同上 | ✅ 同上（既に正しい） |
  - 原因: 上記2件と同根。`evalExpression()` の高速パスは「トップレベル演算子が `=` か」しか見ておらず、`+`/`-` で識別子が左辺（再帰的に、ネストした `+`/`-` チェーンの最も左の葉まで）に来るケースを見ていなかった
  - 修正: 高速パスの判定に、トップレベルが `+`/`-` の `BinaryExpression` で、左辺を `+`/`-` の連鎖を通して再帰的に辿った最終的な葉が `Identifier`/`MemberExpression`（= 呼び出し可能な形）である場合も「statement として解析すべき」に含めるよう拡張。括弧で明示的に囲まれた場合（`(x) + 1`）は曖昧性が解消されるため対象外とする | `eval-statement-ambiguity.test.ts`
  - **既知のスコープ外（意図的に対応しない）**: `eval('Foo(1, 2)')` のように明示括弧付きの関数呼び出しが**単独で** `eval()` に渡された場合、現状は戻り値をそのまま返す（複数文中で同じ呼び出しが行われた場合は statement として戻り値を捨てる、という非対称性が残る）。これは「関数呼び出しの結果をそのまま取得する」という最も基本的で有用な `eval()` の使い方を壊さないよう、意図的に対象外としている
- ✅ **仕様バグ修正（2026-06-27）: `eval()` で `UBound(arr) + 1` / `Len("x") + 1` 等の組み込み関数戻り値への算術演算が `Error 424: Object required` になる**
  - 原因: `isCallableLeftmostLeaf()` が `CallExpression` ノードに対して `true` を返していたため、`UBound(arr) + 1` の左辺（`CallExpression`）が「曖昧な呼び出し文」と誤判定された。fast-path（式評価）がスキップされ、フォールバックの文パーサーが `UBound(arr + 1)`（配列オブジェクトに 1 を加算）として誤解析し、Error 424 が発生していた。括弧ワークアラウンド `(UBound(arr)) + 1` では外括弧でルートノードが `ParenthesizedExpression` になるため fast-path が正しく使われていた
  - 修正: `isCallableLeftmostLeaf()` の `CallExpression` ケースを `return false` に変更。明示括弧付きの関数呼び出し `foo(args)` は括弧で引数リストが確定しており、`foo args+1` のような暗黙 Call 文との曖昧性はない | `eval-builtin-arithmetic.test.ts`
- ✅ **Fix: `Private/Public Type` 宣言が無視される** | `private-type-declaration.test.ts`
  - 原因: パーサーがスコープ修飾子（`Private`/`Public`）の後に `Type` キーワードが来るケースを未処理。`parseDimStatement` にフォールスルーし、`Type` がキーワードトークンのため変数名チェックに失敗してエラーリカバリされていた
  - 症状: `Private Type json_Options ... End Type` が AST に含まれず、`json_Options` 型の変数が常に `0` になる
  - 修正: `Public/Private/Friend` ブロックに `KeywordType` / `KeywordEnum` のケースを追加。`parseTypeDeclaration` の配列メンバー（`name(0 To N) As Type`）も括弧をスキップして対応

- ✅ **Fix: `VBA.vbNull` など `VBA.` モジュール修飾定数アクセスが失敗する** | `vba-module-qualifier.test.ts`
  - 原因: `evaluateMemberExpression` が `VBA.SomeThing` を評価する際、`VBA` 変数が未定義のため `0` になりプロパティアクセスでエラー 424 になっていた
  - 症状: `Case VBA.vbNull` / `Case VBA.vbString` などが実行時エラー 424 で失敗
  - 修正: `possibleModule.toLowerCase() === 'vba'` の場合、プロパティ名を直接 `env.get(propName)` で lookup するフォールバックを追加（`vbNull`, `vbString` 等は既に env に登録済み）

- ✅ **Fix: 再帰関数で `Dim` が外側のスコープを上書きする** | `recursive-dim-scope.test.ts`
  - 原因: `evaluateVariableDeclaration` が `this.env.set(varName, initialValue)` を使っており、`set()` はスコープチェーンを遡って既存の変数を上書きする。再帰呼び出し時に内側の `Dim` が外側の同名変数を初期値でリセットしてしまう
  - 症状: `ConvertToJson` が再帰呼び出しされると、内側の `Dim json_Converted As String` が外側の `json_Converted` を `""` にリセットしてバッファー内容が消失。`{"name":"Alice",...}` が `"":true}` になる
  - 修正: プロシージャ内の変数宣言は `setLocally()` を使用してカレントフレームにのみ変数を作成

- ✅ **Fix: `obj.Method(key) = value` / `Set obj.Item(key) = obj2` / `outer(k1)(k2) = val` 形式の代入が失敗する** | `member-call-assignment.test.ts`, `object-member-assignment.test.ts`
  - 原因: `evaluateAssignmentToVariable` が `CallExpression` の LHS を処理する際、`callee` が `Identifier` の場合のみ対応しており、`callee` が `MemberExpression`（`obj.Item(key)` 形式）の場合にエラー 5 を投げていた
  - 症状: `dict.Item(key) = val` のように `obj.Method(args) = val` 形式で代入すると実行時エラー 5 "Invalid procedure call or argument" が発生
  - 修正: `evaluateAssignmentToVariable` に `callee=MemberExpression`（`obj.Item(k) = v`）と `callee=CallExpression`（`outer(k1)(k2) = v`）のケースを追加
  - `evaluateSetStatement` にも同様の `CallExpression` LHS ケースを追加（`Set obj.Item(k) = obj2`）

- ✅ **Fix: 関数内で `F = F & "X"` が再帰呼び出しになり、戻り値変数の初期値が `0` になる** | `function-return-var.test.ts`
  - 原因1: `Identifier` 評価時に「0引数なら auto-call」するロジックが、現在実行中の関数名（= 戻り値変数）にも適用されていた
  - 原因2: `callProcedure`（公開 API）が `evaluateCallExpression` と異なり、関数の戻り値変数を `vbaEmpty` で初期化していなかった。`env.get` が implicit initialization で `0` を返していた
  - 症状: `F = F & "X"` がスタックオーバーフロー、または戻り値変数が `""` でなく `0` になる
  - 修正: Identifier 評価で `currentProcedureName` と一致する場合は auto-call しない。`callProcedure` に `localEnv.setLocally(proc.name.name, vbaEmpty)` を追加

- ✅ **Fix: `On Error Resume Next` が `Select Case` / `If` / `For` / `While` / `With` ブロック内部で機能しない** | `on-error-in-blocks.test.ts`
  - 原因: `evaluateSelectCaseStatement`・`evaluateIfStatement`・`evaluateForStatement`・`evaluateForEachStatement`・`evaluateDoWhileStatement`・`evaluateWhileStatement`・`evaluateWithStatement` が単純な `for...of` ループでブロック内ステートメントを実行しており、`On Error Resume Next` のエラーキャッチが機能しなかった。エラーはブロック全体を貫通して `executeStatements`（プロシージャ本体レベル）に伝播し、Select Case 全体・If 全体が「1ステートメント」としてスキップされていた
  - 症状: `Select Case VarType(arr)` 内で `LBound(arr, 2)` が VBA エラー 9 を投げると、`On Error Resume Next` で次行に進まず Select Case 全体がスキップされ、戻り値が null になる（`ConvertToJson(array)` が `""` を返す）
  - 修正: 各ブロック評価関数でブロック本体の実行を `executeStatements(body, 0, false)` に変更。`isTopLevel=false` の場合は VBA エラーのみ `On Error Resume Next` で処理し、`GoTo` / `GoSub` / `Return` / `Resume` はプロシージャ本体の `executeStatements` に委譲するよう再スロー

- ✅ **Fix: `VarType(Nothing)` が 9（vbObject）ではなく 12 を返す** | `vartype.test.ts`
  - 原因: `vbaNothing` は `Symbol` であり、`vartype` 組み込み関数の各型チェック（`instanceof`・`typeof object`）をすべてすり抜けて `return 12` にフォールスルーしていた
  - 症状: `JsonConverter.bas` の `json_IsUndefined` が `VarType = vbObject (9)` で Nothing を検出する設計のため、辞書の Nothing 値がスキップされず JSON に出力されてしまう
  - 修正: `vartype` の先頭近くに `if (val === vbaNothing) return 9; // vbObject` を追加

- ✅ **Fix: `Set c = New Collection` / 関数から返した Collection への `For Each` が "Type mismatch" になる** | `for-each.test.ts`
  - 原因: `Dim col As New Collection` は `VbaCollection` クラスインスタンス（`items` getter あり）を生成するが、`Set c = New Collection` は `instantiateClass('Collection')` → `registerExternalObject('Collection', ...)` のファクトリ経由でプレーンオブジェクト（`Symbol.iterator` あり・`items` なし）を生成する。`evaluateForEachStatement` が `items` プロパティの存在のみをチェックしており、ファクトリオブジェクトを処理できなかった
  - 症状: `ParseJson("[1,2,3]")` が返す Collection への `For Each` でエラー 13 が発生
  - 修正: `evaluateForEachStatement` に `__isVbaCollection__` チェックを追加し、`Symbol.iterator` があれば `Array.from` で列挙するパスを設ける

- ✅ **Fix: `Val()` が科学表記（`1.5e2` など）を正しく変換しない** | `conversion-int-val.test.ts`
  - 原因: 正規表現が `/^[+-]?\d*(\.\d*)?/` で `eE` を含まず、`Val("1.5e2")` が `1.5` を返していた
  - 症状: `json_ParseNumber` 内で `VBA.Val("1.5e2")` を呼ぶと 1.5 になり、科学表記の数値が正しくパースされない
  - 修正: 正規表現に `([eE][+-]?\d+)?` を追加

- ✅ **Fix: `vbBack`・`vbFormFeed` 定数が未定義** | `builtin-strings.test.ts`
  - 原因: 定数登録が漏れており、`json_ParseString` の `\b`・`\f` エスケープで `vbBack`/`vbFormFeed` を参照するとエラーになる
  - 修正: `vbBack = "\b"`, `vbFormFeed = "\f"` を定数として登録

- ✅ **Fix: `Hex$`・`Oct$` が未登録（`hex`/`oct` のみあり）** | `number_literals.test.ts`, `builtin-strings.test.ts`
  - 原因: `$` なし版のみ登録しており `Hex$`・`Oct$`・`ChrW$` 呼び出しで Error 438 になっていた
  - 症状: `json_Encode` 内の `"\u" & Right$("0000" & Hex$(AscW), 4)` が失敗し、非ASCII文字の `\uXXXX` エンコードが動かない
  - 修正: `hex$`・`oct$`・`chrw$` を各々のエイリアスとして登録

- ✅ **Fix: `Format$` の `HH`（大文字）が時刻として認識されず、`mm` が文脈判定なしに常に月になっていた** | `format.test.ts`
  - 原因1: `formatDate` の正規表現が小文字 `hh` のみ対応で `HH` をリテラルとして出力していた
  - 原因2: `mm` が月か分かの文脈判定がなく、`"HH:mm:ss"` の `mm` が常に月になっていた
  - 修正: トークン単位の処理に変更し、直前が `h`/`hh` のとき `mm`/`m` を分として解釈する `prevTokenWasHour` フラグを実装

- ✅ **Fix: `Dim ws As Worksheet` など未知の外部オブジェクト型の初期値が `Nothing` でなく `Empty` になる** | `on-error-nothing.test.ts`
  - 原因: 変数宣言の初期化ロジックが `classDefinitions`/`externalObjectFactories` に未登録の型を "Variant や未知の型" と同扱いして `vbaEmpty` を返していた
  - 症状: `Dim ws As Worksheet` で `ws Is Nothing` が False になり、`On Error Resume Next` で `Set` が失敗した後もワークシートアクセスで Error 91 が発生する
  - 修正: `Variant`・`Date` 以外の未知型は外部COMオブジェクト型とみなし `vbaNothing` で初期化

- ✅ **Fix: `Format()` 数値フォーマットの零埋め（`"000"`）が動作せず、また `#` プレースホルダーが `0` と同じ動作になっていた** | `format.test.ts`
  - 原因1: 整数部フォーマットの `0` 個数から最小桁幅を算出して `padStart` する処理が未実装だった
  - 原因2: `#`（有効桁のみ表示）と `0`（常に表示）を区別するロジックがなく、`Format(0, "#")` が `""` でなく `"0"` を返し、`Format(0.5, "#.#")` が `".5"` でなく `"0.5"` を返していた
  - 修正: `minIntegers` を整数部の `0` 個数から算出して `padStart` を追加し、`#` のみのフォーマットで整数部が `"0"` の場合は空文字に抑制するよう修正

- ✅ **Fix: `evaluateDateLiteral` が UTC ベースの Date を生成していたため `formatDate` と時差分ずれていた** | `format.test.ts`
  - 原因: `evaluateDateLiteral` が `Date.UTC(...)` で JS Date を生成していたが、`formatDate` はローカル時刻の `getHours()` 等を使うため時差分ずれていた
  - 修正: `evaluateDateLiteral` を `new Date(y, m-1, d, h, min, s)`（ローカル時刻）に変更し、`Now()`/`DateSerial()` と一貫した挙動にした。`formatDate` はローカル時刻の `get*` メソッドのまま
  - 影響: `ConvertToJson(Date)` の ISO 8601 出力が正しくなった

- ✅ **Fix: `On Error GoTo` ハンドラー内の `If`/`For`/`While` ブロックで `Err.Raise` すると無限ループになる** | `error-handler-reentry.test.ts`
  - 原因: ネストブロックの `executeStatements(isTopLevel=false)` が re-throw 前に `isInErrorHandler = false` にリセットしていたため、外側の `executeStatements` がフラグを見て「ハンドラー外」と誤判断し、同じ GoTo ラベルに再ジャンプしてループ
  - 修正: "bubble up" パスの `this.isInErrorHandler = false` を削除。ネストブロックからの re-throw では `isInErrorHandler` を保持し、外側が正しくバブルアップする
  - 影響: `LibBook.bas` の `NewWorkbook` 関数（Cleanup ハンドラー内で `Err.Raise` するパターン）がハングしなくなった

- ✅ **Fix: `Append`・`Output`・`Binary`・`Random` がファイルモードキーワードとして予約語化され変数名/式に使えない** | `backslash-escape-min.test.ts`
  - 原因: Lexer が `append` / `output` / `binary` / `random` を無条件に `KeywordAppend` 等のトークンに変換していたため、`Dim append As String` の Dim 宣言後の代入（`append = "\\\\"` など）がパースエラーになっていた
  - 症状: `json_BufferAppend` 内の `json_Append` ではなく、ユーザーコードで `append` などを変数名に使うとパース失敗。スタックトレースは "Unexpected token in expression 'append'"
  - 修正: `parseStatementInner` の識別子ステートメント開始条件と `parsePrimary` の文脈的識別子リストに `KeywordOutput`/`KeywordAppend`/`KeywordRandom`/`KeywordBinary` を追加（`Open...For Append` の文法は既に `parseOpenStatement` で正しく処理される）

- ✅ **Fix: `Sub (expr)*x` 形式でスペース前 `(` が関数呼び出しとして貪欲解析され ParseError になる** | `sub-call-paren-expr.test.ts`
  - 原因: `parsePrimary()` のポストフィックスループが `Debug.Print (1+2)` を `CallExpression(Debug.Print, [1+2])` として貪欲消費し、後続の `*3` が新引数の先頭トークンとして解析されようとしたが `*` は式の先頭に置けないため ParseError
  - 症状: `Debug.Print (1+2)*3` が "Unexpected token in expression '*'" でパースエラー。VBA の正しい動作は `9` を表示
  - 修正: `isBinaryOnlyOperator()` で単項不可の二項演算子 (`*`,`/`,`^`,`\`,`+`,`-`,`&`,`Mod`, 比較演算子) を判定。文ステートメントで `parsePrimary()` が `CallExpression` を返し次トークンが該当演算子の場合、位置をリセットして `parsePrimary(stopBeforeSpacedLParen=true)` で再解析。スペース前の `(` をポストフィックス消費せず、`parseCallArgument()` が `(1+2)*3` を一つの引数式として正しく解析する

- ✅ **Fix: For/ForEach/DoWhile/While ループ本体内のラベルへの GoTo がエラーになる** | `goto-in-loop.test.ts`
  - 原因: `executeStatements(body, 0, isTopLevel=false)` は GoTo 例外を無条件で re-throw する設計で、ループ本体（`stmt.body`）内にあるラベルを自分では探さなかった。`evaluateForStatement` 等が GoTo を受け取っても `Exit For` 以外はすべて上位に投げていたため、トップレベルの `executeStatements` が手続き本体からラベルを探し、ループ内にしか存在しないラベルを見つけられずエラーになっていた
  - 修正: `evaluateForStatement` / `evaluateForEachStatement` / `evaluateDoWhileStatement` / `evaluateWhileStatement` の catch ブロックで GoTo を受け取ったとき、`stmt.body` 内にラベルが存在すれば `executeStatements(stmt.body, labelIndex + 1, false)` でラベル直後から残りを実行するよう変更。`findLabelInBody` ヘルパーを追加
  - 影響: MS-VBAL §5.2.1.2 — ラベルのスコープは手続き全体（ループの内外を問わない）という仕様に準拠

- ✅ **Fix: Option Explicit チェックを呼び出し時の env 状態で判定する** | `cross-module-const.test.ts`
  - 原因: `optionExplicitViolations` に記録された未宣言名を呼び出し前に固定判定していたため、`runner.set()` や別モジュールで後から定義した定数が解決済みと扱われなかった
  - 修正: `callProcedure` 内のチェックを `env.hasVariable(name)` で再評価するよう変更。呼び出し時点の env に名前が存在すれば通過する（`optionExplicitViolations` は `Map<procName, Set<undeclaredNames>>` に変更）
  - 詳細: `docs/internals/NAME_RESOLUTION.md` 参照

- ✅ **Fix: モジュールレベル定数のクロスモジュール参照がロード順依存になっていた** | `cross-module-const.test.ts`
  - 原因: `evaluate()` 実行時に定数の右辺を即評価するため、参照先モジュールがまだロードされていない場合に `env.get()` の暗黙初期化（0）が採用されていた
  - 修正: `VBARunner` に二段階ロードを導入。Pass 1（`evaluateModule()`）で全モジュールをロード後、Pass 2（`resolveIdentifiers()`）で全モジュールレベル定数を再評価。依存グラフをトポロジカルソートして正しい順序で確定させる
  - 仕様: VBA では全モジュールがコンパイル時に一括解決されるため、ロード順によらずクロスモジュール定数参照が動作する（MS-VBAL §5.6.10 tier 4）

- ✅ **Fix: `Private Const` の他モジュールからのアクセスをエラーにする** | `cross-module-const.test.ts`
  - 仕様: VBA では `Private Const` は同一モジュール内からのみ参照可。他モジュールから参照すると「コンパイルエラー 定数式が必要です」
  - 実装: `resolveIdentifiers()` の依存グラフ構築時に、参照先 Const の `scope` が `'private'` かつ別モジュールであればエラーを throw

- ✅ **Fix: モジュールレベル定数の循環参照を検出してエラーにする** | `cross-module-const.test.ts`
  - 原因: 二段階評価の Pass 2 を単純に一度だけ実行した場合、相互参照する定数は不定値になりエラーにもならなかった
  - 修正: `topologicalSortConsts()` で DFS による閉路検出を実装。循環が見つかれば `"Circular reference in constant declarations: A → B → A"` 形式のエラーを throw する
  - 仕様: VBA コンパイラは定数の循環参照をコンパイルエラーとして検出する

- ✅ **Fix: `For i ... Next j` のようにループ変数と Next 後の変数が不一致の場合にコンパイルエラーを生成しない** | `for-next-variable-mismatch.test.ts`
  - 原因: パーサーが `nextIdentifier` を AST に記録するだけで、ループ変数（`identifier` / `variable`）との一致検証を行っていなかった
  - 修正: `parseForStatementBody` と `parseForEachStatementBody` の `Next` 解析直後に大文字小文字を無視した名前比較を追加し、不一致なら `throwError` でコンパイルエラーを throw する。`Next`（変数名なし）は引き続き正常
  - 仕様: VBA コンパイラは「Next で指定された変数の参照が不正です」としてコンパイルエラーを生成する
