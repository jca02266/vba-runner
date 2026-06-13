# REPORT.md

## 作業報告

Excel VBA シミュレーションワークスペースを作成し、3つのデモ（RunWorkbookDemo / RunDataDemo / RunValidationDemo）を実行検証した。

---

## 検証結果

| デモ | 状態 | 内容 |
|------|------|------|
| RunWorkbookDemo | ✅ 正常 | Workbook_Open / BeforeClose / SheetChange / SheetActivate イベントチェーン |
| RunDataDemo | ✅ 正常 | AutoSort / CountIf / SumIf / VLookup / Error 91 捕捉 / レポート生成 |
| RunValidationDemo | ✅ 正常 | Worksheet_Change による入力検証・自動修正・日付バリデーション |

---

## 発見したエンジンバグ

### B-7: `Event`/`RaiseEvent` でキーワードをイベント名として使えない (修正済み)

- **現象**: `Public Event Open()` や `RaiseEvent Open` でエラー。`Open` / `Close` 等は VBA キーワードだが Excel VBA ではイベント名として正規に使用される
- **修正**: `parseEventDeclaration` / `parseRaiseEventStatement` に `isNameToken` チェックを追加
- **回避策として変更した設計**: `SimWorkbook.cls` の `Open()` / `Close()` メソッド名を `OpenWorkbook()` / `CloseWorkbook()` に変更（メソッド名としてキーワードを使うとメソッド内の戻り値代入 `Close = value` も失敗するため）

### 情報: VBA キーワードをメソッド名・戻り値代入に使えない制約 (未修正)

- **現象**: `Public Function Close() As Boolean` を宣言し `Close = Not flag` と書くと、パーサーが `Close` を file I/O キーワードとして解釈してエラーになる
- **実際の VBA**: クラスモジュールでは `Function Close()` は合法で、戻り値代入も有効
- **ステータス**: 未修正。設計でキーワードとの衝突を避けることで回避

---

## VBA ソース設計上の注意点

1. **`rem` は予約語（コメント記号）**: `Dim rem As Long` は不正。変数名は `colRem` 等に変更
2. **`Open` / `Close` はキーワード**: メソッド名やイベント名として使う場合は `OpenWorkbook` 等の代替名を推奨
3. **シートイベントのバブリングは手動実装が必要**: `SimSheet.Change` → `SimWorkbook.SheetChange` の伝播は `m_ws1_Change` 内で `m_wb.NotifySheetChange` を明示的に呼ぶことで実現
