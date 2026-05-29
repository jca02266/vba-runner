# Core.bas vs v5.bas 比較考察

## 前提

| ファイル | 位置づけ | 作成経緯 |
|---|---|---|
| `TaskScheduler.bas` + `TaskScheduler_Core.bas` | プロダクションコード（2 ファイル） | 人間の指示でAIが段階的にリファクタリング |
| `TaskScheduler_v5.bas` | 実験コード（1 ファイル） | Claude サブエージェントが v2+v3+v4 の全手法を統合 |

プロダクションコードは I/O シェル（TaskScheduler.bas）とロジックライブラリ（TaskScheduler_Core.bas）の 2 ファイル構成。
v5.bas と公平に比較するには両ファイルを合算する必要がある。

---

## 定量比較（analyzer 実測値）

### ファイル別内訳

| 指標 | TaskScheduler.bas | TaskScheduler_Core.bas | **合算** | v5.bas |
|---|---|---|---|---|
| 総行数 | 204 | 416 | **620** | 591 |
| プロシージャ数 | 1 | 28 | **29** | 24 |
| 最大プロシージャ行数 | 200 ❌ | 78 ⚠️ | **200 ❌** | **42 ⚠️** |
| 最大ネスト深度 | 2 ✅ | 4 ⚠️ | **4 ⚠️** | **3 ⚠️** |
| ❌ 判定数 | 1（行数200） | 2（ByRef） | **3** | **3** |
| 凝集度 HIGH 率 | 0/1 ❌ | 26/28 | **26/29 (90%)** | **23/24 (96%)** |

### 合算での比較

| 指標 | 合算（.bas × 2） | v5.bas |
|---|---|---|
| 総行数 | 620 | **591** |
| プロシージャ数 | 29 | **24** |
| 最大プロシージャ行数 | 200 ❌（TaskScheduler.bas の AutoScheduleTasks） | **42 ⚠️** |
| 最大ネスト深度 | 4 ⚠️ | **3 ⚠️** |
| ❌ 判定数 | 3 | **3**（同数） |
| 凝集度 HIGH 率 | 26/29 (90%) | **23/24 (96%)** |
| マジックナンバー 0.25/0.1 | ❌ 残存 | ✅ 定数化済み |
| テストラッパー | ✅ あり | ❌ なし |
| 孤立関数（refs=0） | **20/29** | 0/24 |
| UDT 数 | **3**（CalendarConfig / TaskConfig / AssigneeConfig） | 2（TaskLayout / AssigneeLayout） |

> **重要な発見**: TaskScheduler.bas の `AutoScheduleTasks` は 200 行 ❌ で凝集度 LOW ❌。
> 2 ファイルに分割しても「I/O シェル」そのものがまだ大きいままである。

---

## 設計思想の根本的な違い

### Core.bas + TaskScheduler.bas は「2 ファイル分割」

Core.bas の全関数が `Public` で refs=0 が 20/28。これは欠陥ではなく、`TaskScheduler.bas` が呼び出し元となる意図的な設計。

```
TaskScheduler.bas（200行 ❌ · I/O シェルだが大きい）
  ├── InitCalendarConfig / InitTaskConfig / InitAssigneeConfig
  ├── GetLastTaskRow / GetLastCalendarCol
  ├── GetTaskRange / GetScheduleRange / ...
  ├── ScanLockedRows
  ├── ScheduleUnlockedTask（Core.bas の 78 行関数を呼ぶ）
  └── UpdateLevelFinish
```

Core.bas のコールグラフは**フラット**で「部品箱」として設計されているが、
TaskScheduler.bas の `AutoScheduleTasks` 自体が 200 行 ❌ と肥大しており、
**分割の効果が I/O 層でもまだ不十分**という状態になっている。

### v5.bas は「スタンドアロン」

単一ファイルで完結。全関数に内部の呼び出し元があり孤立関数ゼロ。
コールグラフは**4 層の深い階層**。

```
AutoScheduleTasks（26行 ✅）
  └── RunScheduleCore
        ├── ReadMetaData / ReadGridData / ReadHolidayData / ReadConfigData
        ├── BuildCapacityDict
        ├── ScanLockedRows → EnsurePersonUsage, AccumulateLockedRow
        └── ScheduleAllTasks（42行）
              ├── ResolveLockedTaskFinish / UpdateLevelFinish
              └── DispatchUnlockedTask
                    ├── CalcTaskStartIdx
                    └── ScheduleOneTask
                          ├── EnsurePersonUsage
                          └── AllocateDays → CalcDailyAlloc
```

---

## 観点別の優劣

### テスト可能性: Core.bas が優位

Core.bas はテストラッパー（`TestFindLockedTaskFinish` / `TestScheduleUnlockedTask`）を持ち、VBA Runner からテスト意図が明示されている。全関数 Public なので、VBA コード内からの再利用・呼び出しも容易。

v5.bas はほぼ全関数が `Private` だが、**VBA Runner の `run()` は `evaluator.callProcedure()` を直接呼ぶため Private アクセスチェックを受けない**（チェックは VBA コード内からのクロスモジュール呼び出し時のみ適用）。そのため VBA Runner 環境ではテストラッパーなしに個々の Private 関数を直接テストできる。

ただし VBA コード内（別モジュール）から直接呼び出すことはできないため、ライブラリとして再利用する場面では依然として Public 化が必要になる。

### 関数サイズ・ネスト: v5.bas が優位

最大関数が Core.bas 78行/ネスト4 に対し v5.bas は 42行/ネスト3。
v5.bas は `DispatchUnlockedTask → ScheduleOneTask → AllocateDays` と段階的に分割しており、Core.bas の `ScheduleUnlockedTask`（78行）相当のロジックを小さく保っている。

### 定数化: v5.bas が優位

| 値 | Core.bas | v5.bas |
|---|---|---|
| 0.25（割り当て単位） | マジックナンバー×3 | `ALLOC_UNIT` |
| 0.1（マイクロタスク最小） | マジックナンバー×2 | `MICRO_ALLOC_MIN` |
| 0.5（親タスク判定閾値） | マジックナンバー×3 | `PARENT_ALLOC_THRESHOLD` |

### UDT の粒度: Core.bas が優位

Core.bas は `CalendarConfig`（カレンダー位置）/ `TaskConfig`（タスク列）/ `AssigneeConfig`（担当者設定）と責務ごとに 3 分割。

v5.bas は `InitTaskLayout` に COL_・ROW_・STR_LOCK_MARK が混在しており粒度が粗い。

### ByRef 出力パラメーター: 両方とも残存

| ファイル | 件数 | 対象 |
|---|---|---|
| Core.bas | 2 件 | FindLockedTaskFinish, ScheduleUnlockedTask |
| v5.bas | 3 件 | ResolveLockedTaskFinish, AllocateDays, ScheduleOneTask |

どちらも UDT 戻り値化が次の改善候補。

---

## 総合評価（合算ベース）

| 観点 | Core + TaskScheduler | v5.bas |
|---|---|---|
| テスト可能性（VBA Runner） | ✅ 優（テストラッパーあり） | ✅ Private でも直接呼び出し可 |
| テスト可能性（VBAコード内） | ✅ Public で再利用容易 | ⚠️ Private のため別モジュールから呼べない |
| 最大関数サイズ | ❌ 200 行（TaskScheduler.bas） | ✅ 優（42 行） |
| I/O シェルの完成度 | ❌ AutoScheduleTasks が 200 行 | ✅ 優（26 行） |
| 定数化 | ❌ マジックナンバー残存 | ✅ 優 |
| UDT の粒度 | ✅ 優（3 型・責務明確） | ⚠️（2 型・混在あり） |
| ファイル完結性 | ⚠️ 2 ファイル管理が必要 | ✅ スタンドアロン |
| 孤立関数（静的解析上） | ⚠️ 20/29（実際は呼ばれている） | ✅ 全関数に呼び出し元 |

---

## 考察

**2 ファイル合算すると v5.bas の優位性が拡大する。**

Core.bas 単体で見ると最大 78 行と悪くないが、実際の I/O シェルである TaskScheduler.bas の `AutoScheduleTasks` が 200 行 ❌ のままであり、分割の恩恵が I/O 層に届いていない。合算での最大関数は 200 行 vs v5.bas の 42 行となり、差が顕著になる。

**「2 ファイルに分けた」こと自体は正しい方向性だが、I/O シェルの整理が未完。**
TaskScheduler.bas の `AutoScheduleTasks` はロジック関数を順番に呼び出すオーケストレーターとして設計されている。しかし各呼び出しを繋ぐ中間変数の宣言・初期化が多く残っており、100 行超が残存している。

**v5.bas の `AutoScheduleTasks`（26 行）との対比が示す差は、「ロジックを関数に出した後、呼び出し元をどれだけスリムにできたか」の違い。**

### 相互にフィードバックすべき点

| Core + TaskScheduler に取り込む | v5.bas に取り込む |
|---|---|
| 0.25 / 0.1 / 0.5 の定数化 | 3 型の UDT 分割（CalendarConfig 等） |
| ScheduleUnlockedTask のさらなる分割 | Public ラッパーによるテスト可能性の確保 |
| TaskScheduler.bas の AutoScheduleTasks を RunScheduleCore 相当に委譲 | — |

**理想形**: v5.bas の「薄い I/O シェル + 定数化 + 細粒度分割」に、Core.bas の「3 型 UDT + テストラッパー + Public ライブラリ設計」を組み合わせたもの。

---

## 定性的評価

### 可読性

#### Core.bas

**良い点:**
- UDT のフィールド名がそのまま「どの列か」を説明している（`COL_LEVEL`, `COL_OFFSET`）。列番号を知らなくても意味が伝わる。
- 小さなヘルパー（`GetMaxDailyLoad`, `GetNumericCellValue` 等）が命名だけで動作を説明しており、実装を読まなくてよい。
- `AccumulateLockedRowUsage` → `ScanLockedRows` → `ScheduleUnlockedTask` と、動詞の命名から「何をする関数か」が連鎖的にわかる。

**悪い点:**
- `' Refactor #1: Extract Base Start Index Calculation` のようなリファクタリング作業記録がコメントとして残っており、将来の読者にはノイズ。
- 英日混在コメント（英語の Logic 説明 + 日本語の補足）が統一されていない。
- フィールド名が `COL_LEVEL`, `ROW_HEADER` と ALL_CAPS 形式で、VBA の `Const` と区別しにくい。
- `CalcDailyAllocation` 内で `0.25` や `0.1` がマジックナンバーのままコメントで補完されており、コメントなしでは意味が不明。
- `ScheduleUnlockedTask` の引数が 12 個。シグネチャ 1 行が非常に長く、呼び出し側での読解コストが高い。

#### v5.bas

**良い点:**
- ファイル冒頭に設計方針（3 層分離・30 行/ネスト 3 の目安）が明記されており、読者が「何を目指したコードか」を最初に把握できる。
- `' -------------------------` のセクション区切りで「I/O 層」「ビジネスロジック層」が視覚的に分離されており、スクロールしながら目的の層を探しやすい。
- `ALLOC_UNIT`, `PARENT_ALLOC_THRESHOLD` など定数名が「何を意味する値か」を説明している。コメントなしで意図が伝わる。
- `AutoScheduleTasks` が 26 行・ネスト 1 で、マクロ全体の流れを 1 画面で把握できる。

**悪い点:**
- `TaskLayoutConfig` に `AssigneeCol`（担当者列=17）と `NameCol`（担当者設定名前列=17）という同じ値を持つフィールドが混在しており、どちらを使うべきか混乱を招く。
- `DispatchUnlockedTask` という名前が直感的でない。`ScheduleUnlockedTask` や `ProcessUnlockedTask` の方が動作を正確に表す。
- `AllocateDays` の引数が 9 個（ByRef 含む）で、Core.bas の問題を引き継いでいる。
- 全関数が `Private` のため「この関数はどこから呼ばれるか」をコールグラフなしで把握しにくい。

---

### 意図の明確さ

**Core.bas** は「Refactor #N」コメントが各関数に付いており、「元のコードのどの部分を抽出したか」という経緯はわかる。ただし将来の保守者には過去の作業記録は不要であり、「なぜこの関数が存在するか」という設計意図は読み取りにくい。

**v5.bas** は設計方針コメントとセクション見出しにより「この関数はビジネスロジック層である」「I/O シェルである」という役割が明示されている。ただし各関数の「なぜ独立した関数か」という粒度の根拠は書かれていない。

---

### 保守性

**変更シナリオ: 「割り当て単位を 0.25 から 0.5 に変更したい」**

- Core.bas: `CalcDailyAllocation` 内の `0.25` を 3 箇所手動で修正が必要。見落としのリスクあり。
- v5.bas: `ALLOC_UNIT = 0.25` を 1 箇所変更するだけ。

**変更シナリオ: 「カレンダー開始列を変更したい」**

- Core.bas: `CalendarConfig.COL_CALENDAR_START` を `InitCalendarConfig` で変更。影響範囲が型に閉じている。明確。
- v5.bas: `TaskLayoutConfig.CalendarStartCol` を `InitTaskLayout` で変更。同様に明確。

**変更シナリオ: 「スケジューリングロジックに新しいビジネスルールを追加したい」**

- Core.bas: `ScheduleUnlockedTask`（78行）を読んで適切な挿入箇所を探す必要がある。サイズが大きいため修正範囲の特定が難しい。
- v5.bas: `AllocateDays`（日次ループ）か `ScheduleOneTask`（1タスク処理）か `DispatchUnlockedTask`（開始日計算）と、ルールの性質に応じて挿入先が明確。

---

### 拡張性

**新しい列を追加したい場合（例: 優先度列を追加）:**

- Core.bas: `TaskConfig` に新フィールドを追加 → `InitTaskConfig` で初期化 → 使う関数に引数追加。型が明確なので追加は容易だが、Public 関数の引数変更が呼び出し元（TaskScheduler.bas）に波及する。
- v5.bas: `TaskLayoutConfig` に追加 → `InitTaskLayout` で初期化 → Private 関数なので波及は内部のみ。変更範囲が小さい。

**新しいルール（例: 特定の担当者のみ週次割り当て）を追加したい場合:**

- Core.bas: `ScheduleUnlockedTask` が 78 行のモノリシックな関数であるため、追加箇所の特定と検証が難しい。
- v5.bas: `AllocateDays`（日次ループ）に局所的に追加できる。他への影響が見通しやすい。

---

### 命名の一貫性

| 対象 | Core.bas | v5.bas | 評価 |
|---|---|---|---|
| UDT フィールド | ALL_CAPS（`COL_LEVEL`）| camelCase（`CalendarStartCol`）| v5.bas が UDT らしい |
| 定数 | なし（マジックナンバー残存） | UPPER_SNAKE_CASE（`ALLOC_UNIT`）| v5.bas が明確 |
| 関数名 | 動詞+名詞（`CalcDailyAllocation`）| 動詞+名詞（`CalcDailyAlloc`）| 両方良好 |
| フェーズ関数 | `ScanLockedRows`, `ScheduleUnlockedTask` | `ScanLockedRows`, `DispatchUnlockedTask` | Core.bas の方が直感的 |

---

### 定性評価まとめ

| 観点 | Core.bas | v5.bas |
|---|---|---|
| 可読性（全体構造） | ⚠️ フラット構造で把握しやすいが層の境界が不明瞭 | ✅ セクション区切りと設計方針コメントで明確 |
| 可読性（個別関数） | ⚠️ 大きい関数と多引数が読解コストを上げる | ✅ 関数が小さく引数も少ない |
| コメントの質 | ⚠️ 作業記録が残り、英日混在 | ✅ 意図と役割を日本語で統一 |
| 保守性（定数変更） | ❌ マジックナンバー散在 | ✅ 定数 1 箇所変更で済む |
| 保守性（ロジック追加） | ⚠️ 大きい関数への挿入が難しい | ✅ 小さい関数への局所的追加が可能 |
| 拡張性 | ⚠️ Public 関数変更が呼び出し元に波及 | ✅ Private のため波及が内部のみ |
| 命名の一貫性 | ⚠️ UDT フィールドが ALL_CAPS で Const と混同しやすい | ✅ 命名規則が統一されている |
| テスト可能性（VBA Runner） | ✅ テストラッパーありで明示的 | ✅ Private でも直接呼び出し可 |
| テスト可能性（VBAコード内） | ✅ Public で再利用・呼び出し容易 | ⚠️ Private のため別モジュールから不可 |
