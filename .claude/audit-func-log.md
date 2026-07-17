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
| Round | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.10.1.7。digits=Null で JS TypeError クラッシュ→Error 13 に修正 |
| Str | 2026-07-17 | ✅ OK | §6.1.2.3.1.20。Str(Null)→error 94 は実VBA準拠（spec誤り）。正値に空白プレフィックスあり |
| Join | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.16。非配列→String化/Null要素→JS TypeError を Error 13 に修正 |
| DateAdd | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.4.1.1。無効Interval文字列でエラーなし→Error 5 を throw |
| Nz | 2026-07-17 | ✅ OK | MS-VBAL外（Access固有）。Null/Empty→valueifnull の動作は適切 |
| DateSerial | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.4.1.4。2桁年 0-29 が 2000-2029 でなく 1900-1929 になる（JS legacy behavior）|
| DateDiff | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.4.1.2。無効Interval文字列で 0 返却→Error 5 を throw |
| DatePart | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.4.1.3。無効Interval文字列で 0 返却→Error 5 を throw |
| Left | 2026-07-17 | ✅ OK | §6.1.2.11.1.19。Null→Null, 負値→Error 5, 超過→全体返却 すべて正常 |
| Right | 2026-07-17 | ✅ OK | §6.1.2.11.1.30。Left と同様、すべて正常 |
| Mid | 2026-07-17 | ✅ OK | §6.1.2.11.1.25。Null→Null, Start<1→Error 5, Length省略→末尾まで すべて正常 |
| Weekday | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.4.1.13。firstdayofweek=Null で Number(Symbol) JS TypeError→Error 5 に修正 |
| MonthName | 2026-07-17 | ✅ OK | §6.1.2.11.1.28。範囲外→Error 5、Null→Null、Abbreviate 正常 |
| WeekdayName | 2026-07-17 | ✅ OK | §6.1.2.11.1.43。範囲外→Error 5、fdow 変換ロジック正常 |
| CInt | 2026-07-17 | ✅ OK | §6.1.2.3.1.7。Null→error 94、overflow→error 6、Banker's round 正常 |

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
| Round | digits=Null で JS TypeError クラッシュ（仕様: Type Mismatch Error 13） | `builtins.test.ts` | ✅ 修正済み |
| Join | 非配列引数で String 化・Null要素で JS TypeError クラッシュ（仕様: Error 13） | `split-join.test.ts` | ✅ 修正済み |
| DateAdd | 無効 Interval 文字列でエラーなし（仕様: Error 5 を raise） | `dateadd-monthend-rollover.test.ts` | ✅ 修正済み |
| DateSerial | 2桁年 0-29 が 2000-2029 でなく 1900-1929（JS legacy behavior: new Date(10,...) → 1910）| `datetime.test.ts` (Bug CU) | ✅ 修正済み |
| Weekday | firstdayofweek=Null で Number(Symbol) JS TypeError クラッシュ（仕様: VBA Error）| `datetime.test.ts` (Bug CV) | ✅ 修正済み |
| DateDiff | 無効 Interval 文字列で 0 返却（仕様: Error 5 を raise） | `dateadd-monthend-rollover.test.ts` (Bug CS) | ✅ 修正済み |
| DatePart | 無効 Interval 文字列で 0 返却（仕様: Error 5 を raise） | `dateadd-monthend-rollover.test.ts` (Bug CT) | ✅ 修正済み |
