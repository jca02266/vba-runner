# TODO_EXT.md — VBA Runner 拡張機能 改善 TODO

<!-- 優先度凡例: 高／中／低  状態: ❌未着手 🚧作業中 ✅完了 -->

## A. 入力補助

### A-1: シグネチャヘルプ ✅ 完了
`(` / `,` でパラメーター名をポップアップ表示。MsgBox・Format・自前 Sub/Function 対象。
- Hover でシグネチャは見えていたが、入力中に自動表示されなかった問題を解消済み。

### A-2: `.` 後のメンバー補完 ❌ 未着手（優先度：低）
`Scripting.Dictionary` / FSO / 自前クラスに `.` を打ったときメンバー一覧を出す。
- **現状**: 補完はスコープ内シンボル＋組み込み関数のフラットリストのみ。オブジェクト型を意識していない。
- **課題**: 型推論が必要で実装コストが高い。

### A-3: スニペット ✅ 完了
`build/extension/snippets/vba.json` を新規作成し `package.json` に登録。

| prefix | 展開内容 |
|--------|---------|
| `fe` | For Each ... In ... Next |
| `for` | For ... To ... Next |
| `sc` | Select Case ... Case Else ... End Select |
| `if` | If ... Then ... Else ... End If |
| `oeg` | On Error GoTo ... ErrHandler パターン |
| `wi` | With ... End With |
| `sub` | Sub ... End Sub |
| `fn` | Function ... End Function |
| `do` | Do While ... Loop |
| `dim` | Dim 変数宣言 |

---

## B. インデント・整形

### B-1: `Select Case` / `Case` のインデント ✅ 完了（3件対応）
- **Format Document**（`Shift+Alt+F`）: `formatter.ts` で `selectDepth` を使い `Case` 行を `Select Case` と同レベルに配置。実装済み。
- **タイピング中の自動デデント**: `language-configuration.json` の `decreaseIndentPattern` に `Case\b` が欠けていたため `Case` を入力しても行が戻らなかった → 追加済み。
- **`Case X:` 後のボディインデント**: `auto-parens.ts` の `needsBodyIndent` が `/^Case\b/i` にマッチ（コロンの有無は無関係）し、Enter 後に次行を1レベル深くインデントする。実装済み。

### B-2: `With` ブロック内の `.` 補完連動 ❌ 未着手（優先度：低）
`With obj` の中で `.` を打ったとき `obj` のメンバーを補完（A-2 の With 版）。A-2 の実装が前提。

### B-3: Format Document コマンドの公開 ✅ 完了
`formatter.ts` 実装済み、`registerDocumentFormattingEditProvider` で `Shift+Alt+F` および右クリック「ドキュメントのフォーマット」として登録済み。
`editor.formatOnSave` は VS Code 標準の仕組みで動作する（`registerDocumentFormattingEditProvider` 登録済みのため、ユーザーが `"[vba]": { "editor.formatOnSave": true }` を設定するだけで有効）。拡張機能側の追加実装は不要。

---

## C. 診断

### C-1: 未使用変数の診断表示 ✅ 完了
- **実装**: `vba-lint.ts` に VBA014 ルールを追加。宣言後に一度も参照されないローカル変数を Warning として波線表示。
- **有効化方法**: `vba-runner.lint.enabled: true` または `vba-runner.lint.enabledCodes: ["VBA014"]` を設定。
- **VBA009との役割分担**: VBA014 は「一切言及なし」、VBA009 は「代入のみで未読取り（dead store）」。

### C-2: `Option Explicit` なし警告 ＋ クイックフィックス ✅ 完了
- **警告**: VBA013 ルール（`vba-lint.ts`）が既に実装済み。`lint.enabled: true` で有効化。
- **クイックフィックス**: VBA013 診断がある行の電球アイコンから「Add 'Option Explicit'」を選択するとファイル先頭に自動挿入。`isPreferred = true` のため Enter で即適用可能。

### C-3: 宣言前使用の警告 ✅ 完了
`Option Explicit` あり環境で未宣言変数を使ったときの診断表示。
- **実装**: `server.ts` の `getDiagnostics()` で `checkOptionExplicit(ast)` を呼び出し。`option-explicit-checker.ts` の既存ロジックを再利用。
- **動作**: `Option Explicit` があるファイルで未宣言変数を使うと赤い波線 (Error) がリアルタイム表示。`Option Explicit` がなければ何も表示しない。

---

## D. ナビゲーション

### D-1: ワークスペースシンボル（`Ctrl+T`） ✅ 完了
`registerWorkspaceSymbolProvider` を登録。`documentMap` の全 VBA ファイルを走査し、`getDocumentSymbols` の結果を `SymbolInformation` に変換して返す。クラスメンバーも `containerName` 付きで表示。

### D-2: アウトラインのセクション区切り認識 ✅ 完了
`' --- Region ---` や `' === Name ===` 形式のコメントをアウトラインに Namespace シンボルとして表示。
- **実装**: `symbol-provider.ts` の `extractSymbols` がオプションでソース文字列を受け取り、行スキャンでセクションヘッダーを検出。純粋なセパレータ行 (`'==========`) はアウトラインに出ない。

---

## E. テスト・デバッグ体験

### E-1: テスト実行結果のインライン表示 ❌ 未着手（優先度：低）
テストが PASS/FAIL したときコードレンズの表示を変える（例: `✓ Tested (last: 3ms)` / `✗ FAILED`）。
- **現状**: 実行後にターミナルを見る必要がある。

### E-2: `Debug.Print` の Output Channel 分離 ✅ 完了
`Debug.Print` 出力を専用 Output Channel "VBA Debug" に分離。
- **実装**: `extension.ts` で `vscode.window.createOutputChannel('VBA Debug')` を作成。`debugProcedure` コマンドと `TestRunner` 両方の `Evaluator` を "VBA Debug" チャンネルへ接続。出力時に `show(true)` で自動表示（フォーカスは奪わない）。
- **TestRunner の変更**: コンストラクターに `onPrint` を追加、`server.ts` に `setDebugPrintHandler()` を追加。

---

## 優先度まとめ

| 優先 | ID | 状態 | 理由 |
|------|----|----|------|
| ✅ | C-1 未使用変数の診断表示 | 完了 | VBA014 ルールを vba-lint.ts に追加 |
| ✅ | A-3 スニペット | 完了 | snippets/vba.json + package.json 登録 |
| ✅ | C-2 Option Explicit クイックフィックス | 完了 | VBA013 診断 + QuickFix 自動挿入 |
| ✅ | D-1 ワークスペースシンボル | 完了 | registerWorkspaceSymbolProvider 登録 |
| 低 | A-2 メンバー補完 | ❌ | 型推論が必要で実装コスト大 |
| ✅ | C-3 宣言前使用の警告 | 完了 | checkOptionExplicit を getDiagnostics に接続 |
| ✅ | D-2 アウトライン区切り | 完了 | symbol-provider でコメント行スキャン |
| 低 | B-2 With 内補完 | ❌ | A-2 が前提 |
| 低 | E-1 テスト結果インライン | ❌ | UI polish |
| ✅ | E-2 Debug.Print 分離 | 完了 | VBA Debug Output Channel に接続 |
| ✅ | A-1 シグネチャヘルプ | 完了 | — |
| ✅ | B-1 Select Case インデント | 完了 | — |
| ✅ | B-3 Format Document 公開 | 完了 | — |
