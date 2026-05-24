---
name: feedback-regression-tests
description: リグレッションテストは必ず run_all_tests.sh と run_vba_tests.sh の両方を実行する
metadata:
  type: feedback
---

実装・修正後のリグレッションテストは以下の2つのスクリプトを順番に実行する:

```bash
./run_all_tests.sh && ./run_vba_tests.sh
```

**Why:** テストには2系統あり、片方だけでは漏れが生じる。
- `run_all_tests.sh` — TypeScript テスト全般（engine/lsp/spec）
- `run_vba_tests.sh` — VBA ファイルベースのテスト（tests/spec/vba/）

**How to apply:** ユーザーから指示されたとき、または自分でテストが必要と判断したとき、常にこのスクリプトを使う。出力が抑制されておりトークン消費が少ない。esbuild + node を直接叩くのは使わない。
