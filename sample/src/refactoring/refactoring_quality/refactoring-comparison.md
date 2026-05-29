# TaskScheduler リファクタリング比較

`TaskScheduler_v1.bas` を段階的にリファクタリングし、各手法の効果を比較する。

| バージョン | 手法 | ファイル |
|---|---|---|
| **v1** | オリジナル（リファクタリング前） | `TaskScheduler_v1.bas` |
| **v2** | ガイドなし・analyzer なし — 目視のみ | `TaskScheduler_v2.bas` |
| **v3** | analyzer 出力のみを根拠 | `TaskScheduler_v3.bas` |
| **v4** | `docs/REFACTORING_GUIDE.md` を参照 | `TaskScheduler_v4.bas` |
| **v5** | v2 + v3 + v4 の全アプローチを統合 | `TaskScheduler_v5.bas` |

すべて Claude サブエージェント(Sonnet 4.6)が担当。各エージェントは他バージョンのコードを参照不可。

---

## 定量指標（analyzer 実測値）

| 指標 | v1 | v2 | v3 | v4 | **v5** |
|---|:---:|:---:|:---:|:---:|:---:|
| 総行数 | 394 | 519 | 369 | 468 | 591 |
| プロシージャ数 | 1 | 18 | 6 | 11 | **24** |
| 最大プロシージャ行数 | 391 ❌ | 85 ⚠️ | 247 ❌ | 123 ❌ | **42 ⚠️** |
| 最大ネスト深度 | 8 ❌ | 5 ❌ | 8 ❌ | 6 ❌ | **3 ⚠️** |
| AutoScheduleTasks 行数 | 391 ❌ | 85 ⚠️ | 247 ❌ | 71 ⚠️ | **26 ✅** |
| AutoScheduleTasks ネスト | 8 ❌ | 2 ✅ | 8 ❌ | 1 ✅ | **1 ✅** |
| 凝集度 HIGH 率 | 0/1 | 15/18 | 5/6 ※ | 9/11 | **23/24 (96%)** |
| ❌ 判定の有無 | あり | あり | あり | あり | **なし** |
| 接頭辞クラスター対処 | ✗ | ✗ | ✗ | ✗ | **✓ (UDT化)** |
| 3 層アーキテクチャ | ✗ | ✗ | ✗ | ✓ | **✓** |

※ `AutoScheduleTasks` が LOW ❌ のため実態は悪い

---

## コールグラフ

### v2（目視のみ）
```
AutoScheduleTasks
  ├── BuildCapacityLimits
  ├── ScanLockedRows → IsLockedRow, EnsurePersonExists, ToDouble
  └── ScheduleAllTasks → CalcBaseStartIdx, FindLockedTaskFinish,
                          ScheduleTask → CalcDailyAlloc, UpdateLevelFinish
```
フラットな 2 層。関数数 18 個。

### v3（analyzer のみ）
```
AutoScheduleTasks（247行 ❌）
  ├── SaveApplicationState / RestoreApplicationState
  ├── ReadRangeToArray
  ├── EnsurePersonUsage
  └── UpdateLevelMaxFinish
```
ほぼ v1 と同じ巨大 Sub が残存。外科的な抽出のみ実施。

### v4（REFACTORING_GUIDE.md）
```
AutoScheduleTasks（71行 ⚠️, nest1 ✅）
  ├── SaveApplicationState / RestoreApplicationState
  ├── ReadRangeToArray × 4
  ├── BuildCapacityLimits
  ├── ScanLockedRows
  └── ScheduleAllTasks（123行 ❌）
        ├── CalcBaseStartIdx / EnsurePersonUsage
        ├── CalcDailyAlloc / UpdateLevelMaxFinish
```
I/O シェル化に成功。ScheduleAllTasks が未完。

### v5（全アプローチ統合）
```
AutoScheduleTasks（26行 ✅, nest1 ✅）   ← 薄い I/O シェル
  └── RunScheduleCore（中継層）
        ├── ReadMetaData / ReadGridData / ReadHolidayData / ReadConfigData
        ├── BuildCapacityDict
        ├── ScanLockedRows → EnsurePersonUsage, AccumulateLockedRow
        └── ScheduleAllTasks（42行 ⚠️, nest2 ✅）
              ├── ResolveLockedTaskFinish / UpdateLevelFinish
              └── DispatchUnlockedTask
                    └── CalcTaskStartIdx
                        ScheduleOneTask → EnsurePersonUsage, AllocateDays
                              └── CalcDailyAlloc
```
4 層の明確な階層。全関数が 42 行以下・ネスト 3 以下。❌ ゼロ。

---

## 各手法の効果と限界

### v2（目視）
- **効果**: 全体俯瞰で「Sub が大きすぎる」と判断 → 積極的に分割（最大 85 行）
- **限界**: 接頭辞クラスターや重複の体系的な検出は困難。設計目標が不明確。

### v3（analyzer）
- **効果**: 重複ブロック・マジックナンバーの具体的な行番号付き指摘に忠実に対処
- **限界**: `行数 391❌` という警告では「どう分割するか」の設計判断を与えられない。

  > **発見**: analyzer の「点の指摘」は外科的修正しか生まない。最初から analyzer だけを頼りにするとかえって悪化する。

### v4（REFACTORING_GUIDE.md）
- **効果**: 「I/O 層 / ビジネスロジック層 / データ」の目標アーキテクチャを持って設計 → AutoScheduleTasks が I/O シェルに
- **限界**: ガイドに細粒度の分割指針がなく ScheduleAllTasks（123 行❌）で止まった。

### v5（全統合）
- **効果**: 3 手法を組み合わせることで❌ ゼロを達成。UDT 化・定数化・重複消去・3 層以上の階層・全関数 42 行以下
- **残課題**: ByRef 出力パラメーター（ResolveLockedTaskFinish 等）、100 × 2 の定数化未対応

---

## 残課題比較

| 残課題 | v2 | v3 | v4 | **v5** |
|---|:---:|:---:|:---:|:---:|
| 接頭辞クラスター未対処 | ❌ | ❌ | ❌ | **✅ UDT 化** |
| ByRef 出力パラメーター | ❌ | ❌ | ❌ | ⚠️ 一部残存 |
| AutoScheduleTasks の肥大 | ⚠️ 85行 | ❌ 247行 | ⚠️ 71行 | **✅ 26行** |
| 最大関数の過大（❌） | ❌ 85行 | ❌ 247行 | ❌ 123行 | **✅ なし** |
| 3 層アーキテクチャ | ✗ | ✗ | ✓ | **✓** |
| マジックナンバー残存 | ⚠️ | ⚠️ | ⚠️ | ⚠️ 100×2 |

---

## まとめ

3 つのアプローチは**補完関係**。v5 はそれらを組み合わせることではじめて ❌ ゼロを達成した。

| 何が得意か | 目視 | analyzer | Guide |
|---|:---:|:---:|:---:|
| 全体を小さく分割する | ✅ | ✗ | △ |
| 重複・マジックナンバーを網羅的に検出 | ✗ | ✅ | ✗ |
| アーキテクチャ目標を与える | ✗ | ✗ | ✅ |
| 残課題を定量的に示す | ✗ | ✅ | ✗ |

**推奨順序**:

```
1. Guide で目標アーキテクチャを確認（I/O 分離・3 層構造）
2. 目視でコードを俯瞰し、大きく分割する
3. analyzer で残課題を定量確認・修正
4. 3 を繰り返してクリーンアップ
```

analyzer は「最初の設計を代替するツール」ではなく「**設計後の確認ツール**」として機能する。
