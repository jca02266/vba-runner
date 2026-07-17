# VBA 組み込み関数 仕様・実装統合監査スキル

`src/engine/builtins.ts` および `src/engine/evaluator.ts` に実装されている VBA 組み込み関数を1件選び、
MS-VBAL 仕様書と実装・テストを照合してバグを探す。試した関数は `.claude/audit-func-log.md` に記録し、同じ評価を繰り返さない。

---

## Step 1: 監査ログを読んで未監査の関数を選ぶ

`.claude/audit-func-log.md` の「監査済み関数」テーブルを読む。

次に `builtins.ts` から全関数名を抽出し、まだログにない関数を洗い出す：

```bash
grep -n "ctx\.reg(" src/engine/builtins.ts | sed "s/.*ctx\.reg('\([^']*\)'.*/\1/" | sort -u
```

ログにない関数の中から **1件** を選ぶ。選定の優先順位：

1. 過去に類似バグが出たカテゴリ（文字列・日付・型変換）
2. 引数が複雑（Optional / ParamArray / オーバーロード）な関数
3. まだ選んでいない関数（アルファベット順などで機械的に）

---

## Step 2: 実装を読む

選んだ関数の実装を `builtins.ts` で読む：

```bash
grep -n "'<関数名>'" src/engine/builtins.ts
```

該当行と周辺の実装ロジックを確認する。必要なら `evaluator.ts` も確認する。

---

## Step 3: MS-VBAL 仕様書で該当セクションを確認

`spec/MS-VBAL-index.txt` で仕様書のセクション番号を調べ、本文を読む：

```bash
grep -i "<関数名>" spec/MS-VBAL-index.txt
# → 行番号を取得後:
sed -n '<LINE>,+120p' spec/MS-VBAL.txt
```

確認観点：
- **引数の型・数・Optional の扱い**
- **戻り値の型・値の範囲**
- **Null / Empty / Missing の扱い**
- **境界値・エラー条件**
- **仕様書に書いてある「例外的な動作」**

---

## Step 4: 既存テストを確認

`tests/spec/` に対応テストがあれば読む：

```bash
ls tests/spec/ | grep -i "<関数名>"
# なければ builtins.test.ts や extra_fns.test.ts も確認
grep -rn "<関数名>" tests/spec/ | head -20
```

---

## Step 5: 差分を分析してバグを特定

仕様書・実装・テストの三者を比較し、以下を確認する：

| チェック項目 | 詳細 |
|---|---|
| 引数チェック | 必須引数が足りたとき何が起きるか |
| 型強制 | 仕様書が定める型変換が正しく行われているか |
| Null 伝播 | Null 入力 → Null 返却かどうか（仕様に従っているか）|
| Empty 扱い | Empty は 0・""・False のどれとして扱うべきか |
| 境界値 | Long/Integer 範囲超えのとき Overflow を出すか |
| 文字列バリアント | `$`付き変種（`Left$`等）の挙動が正しいか |
| エラーコード | 不正入力で正しいエラー番号（5, 6, 13 等）が出るか |

バグが見つかった場合は **Step 6** へ。見つからなかった場合は **Step 7** へ。

---

## Step 6: バグが見つかった場合 — テスト追加 → 修正

### 6a. テストを作成する

`tests/spec/<関数名>-audit.test.ts` にテストを追加（既存ファイルがあればそこに追記）：

```typescript
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

// バグを再現するテスト
```

テストを実行して **失敗を確認**：

```bash
npx tsx tests/spec/<テストファイル>.test.ts
```

### 6b. 修正する

`builtins.ts` または `evaluator.ts` を修正してテストをパスさせる：

```bash
npx tsx tests/spec/<テストファイル>.test.ts
```

### 6c. リグレッションを確認

```bash
npx tsx sample/tests/ts/TaskScheduler_Core.test.ts
```

### 6d. コミット（明示的な指示がある場合のみ）

コミットは **ユーザーが明示的に指示した場合のみ** 行う。

---

## Step 7: 監査ログを更新する

`.claude/audit-func-log.md` を更新する：

**監査済みテーブルに行を追加**：

```markdown
| <関数名> | YYYY-MM-DD | ✅ OK / 🐛 バグあり | <備考> |
```

**バグがあった場合は「バグ発見リスト」にも追記**：

```markdown
| <関数名> | <バグ概要> | `<テストファイル名>` | ✅ 修正済み / ❌ 未修正 |
```

---

## Step 8: 結果を報告する

ユーザーに以下を報告する：

1. 監査した関数名
2. 仕様書のセクション番号
3. バグの有無と内容（あれば）
4. 追加したテストファイル名（あれば）
5. 次に監査すべき候補（ログにない関数から上位3件）
