#!/bin/bash

# VBA テスト自動実行スクリプト
# tests/spec/vba/ ディレクトリにある *_Test.{cls,vba} ファイルを自動検出して実行する
#
# このスクリプトは run-all-vba-tests.test.ts テストランナーを実行する
# _Test で終わるすべてのVBAファイルが自動的に検出・実行される
#
# Usage:
#   ./run_vba_tests.sh      # 簡潔出力
#   ./run_vba_tests.sh -v   # 詳細出力

VERBOSE=0
if [ "$1" = "-v" ] || [ "$1" = "--verbose" ]; then
  VERBOSE=1
fi

RUNNER_FILE="tests/spec/vba/run-all-vba-tests.ts"
VBA_DIR="tests/spec/vba"

if [ ! -f "$RUNNER_FILE" ]; then
  echo "❌ VBA test runner not found: $RUNNER_FILE"
  exit 1
fi

if [ ! -d "$VBA_DIR" ]; then
  echo "❌ $VBA_DIR directory not found"
  exit 1
fi

# --- runner 自動生成 ---
# 命名規則: *Test.bas（ベース名 ≤ 24 文字）→ *Test_runner.bas（自動生成・手動編集不可）
# VBA モジュール名上限 31 文字: len(baseName) + len("_runner"=7) ≤ 31 → baseName ≤ 24
generate_missing_runners() {
  local dir="$1"
  ./node_modules/.bin/tsx test-libs/vba-test-generator.ts --missing "$dir" >&2
}

generate_missing_runners "$VBA_DIR"
for subdir in "$VBA_DIR"/*/; do
  [ -d "$subdir" ] && generate_missing_runners "$subdir"
done

if [ $VERBOSE -eq 0 ]; then
  echo -n "Running VBA tests "
else
  echo "--- Starting VBA tests ---"
fi

# tsx で直接実行（esbuild によるバンドル不要）
if [ $VERBOSE -eq 1 ]; then
  echo "Running VBA tests..."
  echo ""
fi

TEST_OUTPUT=$(./node_modules/.bin/tsx "$RUNNER_FILE" 2>&1)
TEST_STATUS=$?

# 簡潔出力の場合、最終行のサマリーのみ表示
if [ $VERBOSE -eq 1 ]; then
  echo "$TEST_OUTPUT"
elif [ $TEST_STATUS -eq 0 ]; then
  # サマリー行のみ抽出
  echo "$TEST_OUTPUT" | tail -3
else
  # エラー時は全出力
  echo "$TEST_OUTPUT"
fi

echo ""

if [ $TEST_STATUS -eq 0 ]; then
  exit 0
else
  exit 1
fi
