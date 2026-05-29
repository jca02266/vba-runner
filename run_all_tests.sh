#!/bin/bash

# VBA Runner のすべてのテストを一括実行するスクリプト
#  - tests/engine/           : エンジン内部のテスト（Lexer, Parser, AST構造）
#  - tests/lsp/              : LSP（Language Server Protocol）のテスト
#  - tests/spec/             : VBA Runner の仕様テスト
#  - tests/test-libs-tests/  : test-libs/ (VBARunner 等) のテスト
#  - sample/tests/ts/        : サンプル VBA コードのテスト
#
# Usage:
#   ./run_all_tests.sh           # 詳細出力を抑制（進捗表示のみ）、デフォルト並列数 2
#   ./run_all_tests.sh -v        # 詳細出力
#   ./run_all_tests.sh -j N      # N 並列で実行（-j 1 でシリアル実行）
#   ./run_all_tests.sh -v -j 4   # 詳細出力 + 4 並列

VERBOSE=0
PARALLEL=2

while [[ $# -gt 0 ]]; do
  case "$1" in
    -v|--verbose) VERBOSE=1; shift ;;
    -j|--jobs) PARALLEL="$2"; shift 2 ;;
    *) shift ;;
  esac
done

# 作業用一時ディレクトリ（終了時に自動削除）
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT

# セマフォ: N スロットを FIFO + FD 3 で実装
SEMAPHORE="$WORK_DIR/sem"
mkfifo "$SEMAPHORE"
exec 3<>"$SEMAPHORE"
for i in $(seq 1 "$PARALLEL"); do printf '\n' >&3; done

# テストファイル収集
FILES=()
for f in tests/engine/*.test.ts tests/lsp/*.test.ts tests/spec/*.test.ts \
          tests/test-libs-tests/*.test.ts sample/tests/ts/*.test.ts; do
  [[ "$f" == *"run-all-vba-tests"* ]] && continue
  [[ -f "$f" ]] && FILES+=("$f")
done

TOTAL=${#FILES[@]}

if [ "$VERBOSE" -eq 0 ]; then
  echo -n "Running tests ($PARALLEL parallel) "
else
  echo "--- Starting all tests ($PARALLEL parallel) ---"
fi

# 各テストをバックグラウンドで実行
IDX=0
PIDS=()

for f in "${FILES[@]}"; do
  # セマフォ取得（空きスロットが出るまでブロック）
  read -u 3

  IDX=$((IDX + 1))
  RESULT_FILE="$WORK_DIR/result_${IDX}.txt"

  (
    out="${f%.ts}.cjs"

    # esbuild でバンドル
    build_out=$(./node_modules/.bin/esbuild "$f" --bundle --outfile="$out" --platform=node 2>&1)
    if [ $? -ne 0 ]; then
      { echo "FAIL"; echo "$f"; echo "❌ Build failed: $f"; echo "$build_out"; } > "$RESULT_FILE"
      printf '.' >&2
      printf '\n' >&3  # セマフォ解放
      exit 0
    fi

    # テスト実行（30秒タイムアウト）
    test_out=$(timeout 30 node "$out" 2>&1)
    test_status=$?

    if [ "$test_status" -eq 124 ]; then
      { echo "FAIL"; echo "$f"; echo "❌ Test timed out (>30s): $f"; } > "$RESULT_FILE"
    elif [ "$test_status" -ne 0 ]; then
      { echo "FAIL"; echo "$f"; echo "❌ Test failed: $f"; echo "$test_out"; } > "$RESULT_FILE"
    else
      { echo "PASS"; echo "$f"; echo "$test_out"; } > "$RESULT_FILE"
    fi

    printf '.' >&2   # 進捗ドット（シングルバイトなのでアトミック）
    printf '\n' >&3  # セマフォ解放
  ) &

  PIDS+=($!)
done

# 全ジョブ完了待ち
for pid in "${PIDS[@]}"; do
  wait "$pid"
done

echo "" >&2

# 結果を投入順に集約（出力のシリアライズ）
TESTS_FAILED=0
FAILED_OUTPUT=""

for i in $(seq 1 "$IDX"); do
  RESULT_FILE="$WORK_DIR/result_${i}.txt"
  [[ -f "$RESULT_FILE" ]] || continue

  status=$(head -1 "$RESULT_FILE")
  fname=$(sed -n '2p' "$RESULT_FILE")
  body=$(tail -n +3 "$RESULT_FILE")

  if [ "$status" = "FAIL" ]; then
    TESTS_FAILED=1
    FAILED_OUTPUT="${FAILED_OUTPUT}${body}\n"
  elif [ "$VERBOSE" -eq 1 ]; then
    echo "--- $fname ---"
    echo "$body"
  fi
done

# VBA テスト（run_vba_tests.sh）を追加実行
if [ -f "./run_vba_tests.sh" ]; then
  if [ "$VERBOSE" -eq 1 ]; then
    ./run_vba_tests.sh -v
  else
    ./run_vba_tests.sh
  fi
  [ $? -ne 0 ] && TESTS_FAILED=1
fi

if [ "$TESTS_FAILED" -eq 0 ]; then
  echo "--- All tests completed successfully! ---"
  exit 0
else
  echo "--- Some tests FAILED ---"
  echo ""
  echo -e "$FAILED_OUTPUT"
  exit 1
fi
