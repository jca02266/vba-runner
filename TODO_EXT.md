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

### A-3: スニペット ❌ 未着手（優先度：中）
`For Each` / `Select Case` / `On Error GoTo` / `With` ブロックなど定型構造を `fe` / `sc` 等で展開。
- **現状**: 未実装。
- **価値**: VBE の補完と違い Tab ストップで変数名・型を順に埋められる。VBE ユーザーの移行コストを下げる。

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

### C-2: `Option Explicit` なし警告 ＋ クイックフィックス ❌ 未着手（優先度：中）
`Option Explicit` が先頭にない `.bas`/`.cls` に警告を出し、クイックフィックスで自動追加。
- **現状**: 未実装。
- **価値**: VBA のベストプラクティス促進。VBE ではデフォルト設定だが VS Code では明示が必要。

### C-3: 宣言前使用の警告 ❌ 未着手（優先度：低）
`Option Explicit` あり環境で未宣言変数を使ったときの診断表示。C-2 が前提。
- **現状**: コンパイルエラーとして実行時に検出されるが、診断として常時表示されているか要確認。

---

## D. ナビゲーション

### D-1: ワークスペースシンボル（`Ctrl+T`） ❌ 未着手（優先度：中）
複数 `.bas`/`.cls` をまたいだシンボル検索。
- **現状**: ドキュメントシンボルはファイル内のみ。大きなプロジェクトで `Sub FindMe` を探す際に不便。

### D-2: アウトラインのセクション区切り認識 ❌ 未着手（優先度：低）
`'===========` や `' --- Region ---` 形式のコメントをアウトラインに折り畳み可能なセクションとして表示。
- VBE でよく使われるコーディング慣習に対応。

---

## E. テスト・デバッグ体験

### E-1: テスト実行結果のインライン表示 ❌ 未着手（優先度：低）
テストが PASS/FAIL したときコードレンズの表示を変える（例: `✓ Tested (last: 3ms)` / `✗ FAILED`）。
- **現状**: 実行後にターミナルを見る必要がある。

### E-2: `Debug.Print` の Output Channel 分離 ❌ 未着手（優先度：低）
`Debug.Print` 出力を専用 Output Channel に分けてフィルタリングしやすくする。
- **現状**: `console.log` に混在。`onPrint` API はあるが拡張機能 UI として露出しているか要確認。

---

## 優先度まとめ

| 優先 | ID | 状態 | 理由 |
|------|----|----|------|
| ✅ | C-1 未使用変数の診断表示 | 完了 | VBA014 ルールを vba-lint.ts に追加 |
| 中 | A-3 スニペット | ❌ | 実装コスト低・VBE ユーザー移行コスト削減 |
| 中 | C-2 Option Explicit クイックフィックス | ❌ | VBA ベストプラクティス促進 |
| 中 | D-1 ワークスペースシンボル | ❌ | 複数ファイル構成で便利 |
| 低 | A-2 メンバー補完 | ❌ | 型推論が必要で実装コスト大 |
| 低 | C-3 宣言前使用の警告 | ❌ | C-2 が前提 |
| 低 | D-2 アウトライン区切り | ❌ | 特定スタイル限定で恩恵が限られる |
| 低 | B-2 With 内補完 | ❌ | A-2 が前提 |
| 低 | E-1 テスト結果インライン | ❌ | UI polish |
| 低 | E-2 Debug.Print 分離 | ❌ | 影響範囲が小さい |
| ✅ | A-1 シグネチャヘルプ | 完了 | — |
| ✅ | B-1 Select Case インデント | 完了 | — |
| ✅ | B-3 Format Document 公開 | 完了 | — |
