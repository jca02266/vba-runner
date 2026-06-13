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

### 情報: `Function` 戻り値代入でキーワード名が使えない制約 (未修正)

- **現象**: `Public Function Close() As Boolean` 内で `Close = Not flag` と書くと、パーサーが `Close` を file I/O キーワードとして解釈してエラーになる
- **実際の VBA**: クラスモジュールでは `Function Close()` は合法で、戻り値代入も有効
- **`Sub` としての宣言は問題なし**: `Public Sub Open()` はコンテキストキーワードとして正常に宣言・呼び出し可能。`wb.Open` も動作する
- **ステータス**: 未修正。`CloseWorkbook` 等の代替名で回避

---

## VBA ソース設計上の注意点

1. **`rem` は予約語（コメント記号）**: `Dim rem As Long` は不正。変数名は `colRem` 等に変更
2. **`Close` の戻り値代入のみ非対応**: `Sub Open()` はコンテキストキーワードとして正常動作。`Function Close()` 内での `Close = value` のみ回避が必要
3. **シートイベントのバブリングは自動**: `SimSheet.SetParentWorkbook` により `SetValue` → `Workbook_SheetChange` の伝播が自動化されている
