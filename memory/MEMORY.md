# Memory Index

- [Node.js ローカルコマンドの実行方法](feedback_node_commands.md) — esbuild 等は `./node_modules/.bin/` で実行、npx は使わない
- [Markdownドキュメント検索は /markdown-query を使う](feedback_use_markdown_query.md) — grep/find でなく mdq で Context を節約
- [Phase 1: MS-VBAL Listed Features Done, Runtime Behaviors In Progress](project_phase_completion.md) — Listed features complete; runtime behaviors (Auto-Instantiation, Default Property, etc.) tracked separately in TODO.md
- [VBA Testing & Refactoring Documentation Architecture](docs_architecture.md) — 4-document system: strategy → refactoring → frameworks → mocks
- [Default Property / Default Member Implementation (13/15 tests)](default_property_implementation.md) — Implicit property access feature; implicit Value assignment working, ByRef bug and Value getter pending
- [Circular Reference & Class_Terminate Lifecycle](runtime_behavior_circular_ref.md) — Class_Terminate invoked once per object; Set obj.property supported
- [TODO items must include test file references](feedback_todo_test_refs.md) — Always append `| \`testfile.test.ts\`` when marking items ✅
- [計画ファイルは実装完了直後に削除する](feedback_plan_cleanup.md) — 残ると毎セッションで無駄にロードされトークンを消費する
- [コミットは明示指示があるまで行わない](feedback_no_auto_commit.md) — 勝手なコミット・プッシュ禁止
- [テストファイルの配置ディレクトリルール](feedback_test_directories.md) — test-libs のテストは `tests/test-libs-tests/`、エンジンテストは `tests/spec/`
- [カタカナ語尾の長音符ルール](feedback_katakana_longvowel.md) — 新規 .md では -er/-or/-ar 語尾に ー を付ける（パラメーター、エディター等）
- [仕様バグ修正時のセット作業](feedback_bug_fix_process.md) — テスト追加 + TODO.md 更新を毎回自発的に行う
- [mdq get の出力フォーマット](feedback_mdq_get_format.md) — `mdq get` は必ず `| python3 -c "import json,sys; print(json.load(sys.stdin)['text'])"` を付けて読む
- [パーサーエラーは diagnostics に記録される](feedback_parser_diagnostics.md) — try-catch では検出不可。ast.diagnostics と ast.body の両方を確認する
- [リグレッションテストは2スクリプト実行](feedback_regression_tests.md) — `./run_all_tests.sh && ./run_vba_tests.sh` を毎回実行する
- [TODO残件確認のルール](feedback_todo_listing.md) — TODO.md と TODO_NEXT.md の両方を確認し、優先度高・難易度低の順で提案する
