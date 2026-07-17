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
| CDate | 2026-07-17 | ✅ OK | §6.1.2.3.1.4。Null→error 94、Empty→error 13、VbaDate→as-is、数値/文字列→変換 正常 |
| CBool | 2026-07-17 | ✅ OK | §6.1.2.3.1.1。Null→error 94、0→False/非0→True、"True"/"False"文字列→正常 |
| CStr | 2026-07-17 | ✅ OK (意図的逸脱) | §6.1.2.3.1.12。CStr(Null)→"" は Bug BK での意図的設計決定（仕様はerror 94だが使い勝手優先）。その他正常 |
| Asc | 2026-07-17 | ✅ OK | §6.1.2.11.1.2。Null→Null、空文字→error 5、AscW と同じ実装 |
| Chr | 2026-07-17 | ✅ OK | §6.1.2.11.1.5。Null→Null、0-255範囲外→error 5、ChrW は-32768〜65535対応 |
| Hex | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.3.1.17。負数 -32767〜-1 が 32ビット（8文字）を返す（仕様: 16ビット4文字）|
| Oct | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.3.1.19。負数 -32767〜-1 が 32ビット（11文字）を返す（仕様: 16ビット6文字）|
| StrReverse | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.40。Null→vbaNull返却（仕様: エラーを発生）→Error 94 に修正 |
| Fix | 2026-07-17 | ✅ OK | §6.1.2.3.1.16。Null→Null、n≥0→floor、n<0→ceil（ゼロ方向への丸め）正常 |
| Int | 2026-07-17 | ✅ OK | §6.1.2.3.1.18。Null→Null、Math.floor（負の無限大方向への丸め）正常 |
| IsArray | 2026-07-17 | ✅ OK | §6.1.2.7.1.2。Array.isArray で正常判定。型付き配列も JS Array なので True ✓ |
| IsNumeric | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.7.1.8。VbaDate を数値型と誤判定（True）→ False に修正。Boolean/数値/Decimal/Currency は True、文字列は数値変換試行 |
| IsDate | 2026-07-17 | ⚠️ 制限あり | §6.1.2.7.1.3。ISO 8601 形式 "2023-01-01" を True と返す（VBA は False）。Date.parse() の過剰受容。VBA日付パーサー実装は複雑のため注記のみ |
| VarType | 2026-07-17 | ⚠️ 制限あり | §6.1.1.16。型付き配列は要素型不明のため常に 8204 (Variant()) を返す（例: Dim a() As String → 8200 が正確）。アーキテクチャ的制限 |
| TypeName | 2026-07-17 | ✅ OK | §6.1.2.7.1.12。VarType と同じ特殊評価パスを使いASTレベルで宣言型を解決。型付き変数は正確に返す |

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
| Hex | -32767〜-1 の範囲で32ビット（8文字）を返却（仕様: 16ビット4文字）| `conversion-int-val.test.ts` (Bug CX) | ✅ 修正済み |
| Oct | -32767〜-1 の範囲で32ビット（11文字）を返却（仕様: 16ビット6文字）| `conversion-int-val.test.ts` (Bug CY) | ✅ 修正済み |
| StrReverse | Null 入力で vbaNull 返却（仕様: エラー発生）→ Error 94 に修正 | `strreverse.test.ts` (Bug CZ) | ✅ 修正済み |
| Abs | 2026-07-17 | ✅ OK | §6.1.2.10.1.1。Null→Null、負→正 正常 |
| Len | 2026-07-17 | ✅ OK | §6.1.2.11.1.22。Null→Null、数値→String 変換後カウント、Empty→0 正常 |
| LBound | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.2.1.2。dim=Null で JS TypeError クラッシュ（Bug DD）、dim=0 で Error 9 未発生（Bug DE）|
| UBound | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.2.1.3。LBound と同じく dim=Null/dim=0 でバグ (Bug DD/DE)  |
| Year | 2026-07-17 | ✅ OK | §6.1.2.4.1.14。Null→Null、日付文字列→変換 正常 |
| Month | 2026-07-17 | ✅ OK | §6.1.2.4.1.9。Null→Null 正常 |
| Day | 2026-07-17 | ✅ OK | §6.1.2.4.1.6。Null→Null 正常 |
| Hour | 2026-07-17 | ✅ OK | §6.1.2.4.1.7。Null→Null、小数部→時間変換 正常 |
| Minute | 2026-07-17 | ✅ OK | §6.1.2.4.1.8。Null→Null 正常 |
| Second | 2026-07-17 | ✅ OK | §6.1.2.4.1.10。Null→Null 正常 |
| StrConv | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.37。conv=Null で JS TypeError クラッシュ → Error 94 に修正。s=Null→Null 伝播は正常 |
| Environ | 2026-07-17 | ✅ OK | §6.1.2.8.1.6。Sandbox経由でOS環境変数を隔離。文字列返却 正常 |
| DateValue | 2026-07-17 | ✅ OK | §6.1.2.4.1.5。Null→Null、日付文字列→VbaDate（時刻部分なし）正常 |
| TimeValue | 2026-07-17 | ✅ OK | §6.1.2.4.1.12。Null→Null、日付文字列→VbaDate（日付部分なし）正常 |
| TimeSerial | 2026-07-17 | ✅ OK | §6.1.2.4.1.11。Null→Null（Bug AX/AY 意図的決定）、時刻オーバーフロー正常（JSのDate巻き上げ活用）|
| Sgn | 2026-07-17 | ✅ OK | §6.1.2.10.1.8。Null→Null（Variant引数）、正/負/0 → 1/-1/0 正常 |
| LCase | 2026-07-17 | ✅ OK | §6.1.2.11.1.17。Null→Null、大文字 → 小文字変換正常 |
| UCase | 2026-07-17 | ✅ OK | §6.1.2.11.1.41。Null→Null、小文字 → 大文字変換正常 |
| LTrim | 2026-07-17 | ✅ OK | §6.1.2.11.1.23。Null→Null、前後空白削除正常 |
| RTrim | 2026-07-17 | ✅ OK | §6.1.2.11.1.23。Null→Null、後空白削除正常 |
| Trim | 2026-07-17 | ✅ OK | §6.1.2.11.1.23。Null→Null、前後空白削除正常 |
| Space | 2026-07-17 | ✅ OK | §6.1.2.11.1.33。負数→Error 5、Null→Null（spec 明記なし）、0→"" 正常 |
| String | 2026-07-17 | 🐛 バグあり (修正済み) | §6.1.2.11.1.38。数値 > 255 で Mod 256 未適用（Bug DB）、空文字 Character で Error 5 未発生（Bug DC）|
| IsNumeric | VbaDate を数値型と誤判定 → True 返却（仕様: §6.1.2.7.1.8 に Date なし → False） | `builtins.test.ts` (Bug DA) | ✅ 修正済み |
| String | 数値 > 255 で Mod 256 未適用（例: String(3,257)→"āāā"、仕様: chr(1)*3） | `builtin-strings.test.ts` (Bug DB) | ✅ 修正済み |
| String | 空文字 Character で Error 5 未発生（String(3,"")→""、仕様: Error 5） | `builtin-strings.test.ts` (Bug DC) | ✅ 修正済み |
| LBound | dim=Null で JS TypeError クラッシュ（仕様: VBA Error） | `builtins.test.ts` (Bug DD) | ✅ 修正済み |
| UBound | dim=Null で JS TypeError クラッシュ（仕様: VBA Error） | `builtins.test.ts` (Bug DD) | ✅ 修正済み |
| LBound | dim=0 で Error 9 未発生（次元は1ベース、dim<1 は無効） | `builtins.test.ts` (Bug DE) | ✅ 修正済み |
| UBound | dim=0 で Error 9 未発生（次元は1ベース、dim<1 は無効） | `builtins.test.ts` (Bug DE) | ✅ 修正済み |
| StrConv | conv=Null で JS TypeError クラッシュ（仕様: VbStrConv 型のため Null 不可 → Error 94）| `strconv_null.test.ts` (Bug DG) | ✅ 修正済み |
