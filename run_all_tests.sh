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
#
# tests/tooling/evaluation-*.test.ts は評価記録など共有リポジトリ内の
# ファイルを変更するため、ほかのテストと並列に実行せず、全並列ジョブの
# 完了後に直列実行する。

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
# *.test.skip.ts は収集対象外（glob *.test.ts にマッチしないため自動的に除外される）
FILES=()
for f in tests/engine/*.test.ts tests/lsp/*.test.ts tests/spec/*.test.ts \
          tests/test-libs-tests/*.test.ts tests/tooling/*.test.ts \
          tests/extractor/*.test.ts sample/tests/ts/*.test.ts; do
  [[ "$f" == *"run-all-vba-tests"* ]] && continue
  [[ -f "$f" ]] && FILES+=("$f")
done

# スキップファイルの一覧表示（*.test.skip.ts）
SKIP_FILES=()
for f in tests/engine/*.test.skip.ts tests/lsp/*.test.skip.ts tests/spec/*.test.skip.ts \
          tests/test-libs-tests/*.test.skip.ts tests/tooling/*.test.skip.ts \
          sample/tests/ts/*.test.skip.ts; do
  [[ -f "$f" ]] && SKIP_FILES+=("$f")
done
if [ "${#SKIP_FILES[@]}" -gt 0 ]; then
  echo "Skipping ${#SKIP_FILES[@]} test(s): ${SKIP_FILES[*]}"
fi

TOTAL=${#FILES[@]}

# Evaluation tooling tests intentionally mutate files under evaluation/. Keep
# that shared-state group serialized so tests cannot observe each other's
# temporary records. Other tooling tests and test groups retain parallelism.
SERIAL_INDICES=()
for i in "${!FILES[@]}"; do
  if [[ "${FILES[$i]}" == tests/tooling/evaluation-*.test.ts ]]; then
    SERIAL_INDICES+=("$((i + 1))")
  fi
done
SERIAL_COUNT=${#SERIAL_INDICES[@]}

if [ "$VERBOSE" -eq 0 ]; then
  echo -n "Running tests ($PARALLEL parallel, $SERIAL_COUNT tooling serial) "
else
  echo "--- Starting all tests ($PARALLEL parallel, $SERIAL_COUNT tooling serial) ---"
fi

# Run one test and write a result record consumed by the reporting pass below.
run_one_test() {
  local f="$1"
  local result_file="$2"
  local timeout_seconds="${3:-30}"
  local test_out
  local test_status

  # tsx で直接実行（esbuild によるバンドル不要）
  # テスト実行（30秒タイムアウト）
  test_out=$(timeout "$timeout_seconds" ./node_modules/.bin/tsx "$f" 2>&1)
  test_status=$?

  if [ "$test_status" -eq 124 ]; then
    { echo "FAIL"; echo "$f"; echo "❌ Test timed out (>${timeout_seconds}s): $f"; } > "$result_file"
  elif [ "$test_status" -ne 0 ]; then
    { echo "FAIL"; echo "$f"; echo "❌ Test failed: $f"; echo "$test_out"; } > "$result_file"
  else
    { echo "PASS"; echo "$f"; echo "$test_out"; } > "$result_file"
  fi
}

# Parallel tests run first. Tooling tests are deliberately held back until
# every parallel test has finished, so no process can race on evaluation/.
IDX=0
PIDS=()

for f in "${FILES[@]}"; do
  IDX=$((IDX + 1))
  if [[ "$f" == tests/tooling/evaluation-*.test.ts ]]; then
    continue
  fi

  # セマフォ取得（空きスロットが出るまでブロック）
  read -u 3

  RESULT_FILE="$WORK_DIR/result_${IDX}.txt"

  (
    run_one_test "$f" "$RESULT_FILE"
    printf '.' >&2   # 進捗ドット（シングルバイトなのでアトミック）
    printf '\n' >&3  # セマフォ解放
  ) &

  PIDS+=($!)
done

# 全ジョブ完了待ち
for pid in "${PIDS[@]}"; do
  wait "$pid"
done

# Execute shared-state tooling tests in their original collection order and
# only after all parallel jobs have released their resources.
for index in "${SERIAL_INDICES[@]}"; do
  f="${FILES[$((index - 1))]}"
  RESULT_FILE="$WORK_DIR/result_${index}.txt"
  run_one_test "$f" "$RESULT_FILE" 120
  printf '.' >&2
done

echo "" >&2

# 結果を投入順に集約（出力のシリアライズ）
TESTS_FAILED=0
TS_TESTS_PASSED=0
TS_TESTS_FAILED=0
FAILED_OUTPUT=""

for i in $(seq 1 "$IDX"); do
  RESULT_FILE="$WORK_DIR/result_${i}.txt"
  [[ -f "$RESULT_FILE" ]] || continue

  status=$(head -1 "$RESULT_FILE")
  fname=$(sed -n '2p' "$RESULT_FILE")
  body=$(tail -n +3 "$RESULT_FILE")

  if [ "$status" = "FAIL" ]; then
    TESTS_FAILED=1
    TS_TESTS_FAILED=$((TS_TESTS_FAILED + 1))
    FAILED_OUTPUT="${FAILED_OUTPUT}${body}\n"
  else
    TS_TESTS_PASSED=$((TS_TESTS_PASSED + 1))
    if [ "$VERBOSE" -eq 1 ]; then
      echo "--- $fname ---"
      echo "$body"
    fi
  fi
done

TS_TESTS_UNEXECUTED=$((TOTAL - TS_TESTS_PASSED - TS_TESTS_FAILED))
echo "TypeScript test files: ${TOTAL} total, ${TS_TESTS_PASSED} passed, ${TS_TESTS_FAILED} failed, ${TS_TESTS_UNEXECUTED} unexecuted"

# VBA テスト（run_vba_tests.sh）を追加実行
if [ -f "./tests/vba/run_vba_tests.sh" ]; then
  if [ "$VERBOSE" -eq 1 ]; then
    ./tests/vba/run_vba_tests.sh -v
  else
    ./tests/vba/run_vba_tests.sh
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
