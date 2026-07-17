# VBA 組み込み関数 監査ログ

関数仕様（MS-VBAL）と実装（`builtins.ts` / `evaluator.ts`）の照合結果を記録する。
バグが見つかった場合は `tests/spec/` にテストを追加し、TODO_SPEC.md に記録する。

## 監査済み関数

| 関数名 | 監査日 | 結果 | 備考 |
|--------|--------|------|------|
| Replace | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.29 |
| InStr | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.14 |
| InStrRev | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.15 |
| Split | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.35 |
| Filter | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.7。Null Match→error 13 は Bug BE の既存決定を維持。多次元→error 13 は実VBAの正確なエラー番号未確認（5 の可能性あり） |
| Val | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.3.1.21。非文字列引数（数値等）を 0 返却→String強制変換に修正 |
| StrComp | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.36。Compare=Null でエラーにならず誤結果→error 94 に修正 |
| Switch | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.8.1.16。奇数個引数（不対ペア）で undefined 返却→error 5 に修正 |
| Choose | 2026-07-17 | ✅ OK | §6.1.2.8.1.2。Null/範囲外→Null は正常。Index型丸め(Single→Integer変換)は許容範囲 |
| IIf | 2026-07-17 | ✅ OK | §6.1.2.8.1.11。両辺評価はJS関数呼び出しで自動実現。Null条件→error 94 は適切 |

## バグ発見リスト

| 関数名 | バグ概要 | テストファイル | 修正済み |
|--------|---------|--------------|---------|
| Replace | `Replace(Null, "a", "b")` が Null を返す（仕様: error 94）; `Start < 1` が error 5 を出さず誤結果を返す | `builtin-strings.test.ts` (Bug BG) | ✅ 修正済み |
| InStr | `typeof args[0] === 'number'` 判定が VbaBoolean/vbaNull を2引数形式に誤分類; Start=Null・Compare=Null でエラーなし | `builtin-strings.test.ts` (Bug BH) | ✅ 修正済み |
| InStrRev | Start=Null/Compare=Null が JS TypeError かエラーなし（仕様: error）; Start=0/-2 が error 5 でなく誤結果; `InStrRev("", "", 5)`→5（仕様: 空StringCheck→0）; `InStrRev("abc", "", 5)`→5（仕様: Start>Len→0） | `instrrev.test.ts` (Bug CJ) | ✅ 修正済み |
| Split | `Split("", ",")`→1要素`[""]`（仕様: 空配列）; `Split("abc", "")`→文字分割（仕様: 全体1要素）; Compare 引数完全無視（vbTextCompare 不動作）; Delimiter/Limit/Compare=Null が JS TypeError または Symbol 文字列化 | `split-join.test.ts` (Bug CK) | ✅ 修正済み |
| Filter | 多次元配列でエラーにならず誤結果（仕様: "not a one-dimensional array" はエラー） | `filter.test.ts` (Bug CL) | ✅ 修正済み |
| Val | 非文字列引数（42, 3.14 等）を常に 0 返却（仕様: String 強制変換してから解析） | `conversion-int-val.test.ts` | ✅ 修正済み |
| StrComp | `StrComp(s1, s2, Null)` でエラーにならず binary 比較結果を返却（仕様: error 94） | `strcomp.test.ts` | ✅ 修正済み |
| Switch | 奇数個引数（不対ペア）で undefined を返却（仕様: 実行時エラー Error 5） | `choose_switch.test.ts` | ✅ 修正済み |
