# CompileError.bas の CASE 名・Sub 名を補完するスキル

選択中の `CompileError.bas` の CASE ブロックを読み、
`CASE:` と `Sub test()` を適切な名前で書き換える。

## 手順

### Step 1: 選択範囲を確認

IDE の選択テキスト（`ide_selection`）から対象ブロックを特定する。
ブロックは以下の構造を持つ：

```
' CASE:           ← 空欄（これを埋める）
' TYPE: parse|resolve|preproc
' VBA: ...        ← VBE が表示するエラーメッセージ
' RUNNER:
' NOTE: ...       ← 省略可。挙動の補足説明
Sub test()        ← Sub 名（これを埋める）
    ...  ' @error
End Sub
```

### Step 2: CASE 名を決める

以下の優先順で命名する：

1. **エラーの原因となるコード構造**を動詞＋名詞で表す（例: `sub_call_arg_count_mismatch`）
2. VBA エラーメッセージの意味をそのまま英語に（例: `duplicate_dim`）
3. NOTE に記載されている特殊な挙動があればそれを反映

ルール：
- スネークケース（小文字、アンダースコア区切り）
- 既存の CASE 名と重複しないこと（ファイル内を確認）
- 20 文字以内を目安

### Step 3: ファイルを書き換える

- `CASE:` → `CASE: {決めた名前}`
- `Sub test()` → `Sub Case_{決めた名前}()`

Edit ツールを使って該当箇所のみ書き換える。
ブロック内の他の行（TYPE/VBA/RUNNER/NOTE/ボディ）は変更しない。

### Step 4: 既存 CASE 名の重複チェック

```bash
grep "^' CASE:" tests/vba/CompileError.bas
```

重複があれば連番サフィックス（`_2` など）を付ける。
