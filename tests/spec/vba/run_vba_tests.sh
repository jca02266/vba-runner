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

if [ ! -f "$RUNNER_FILE" ]; then
  echo "❌ VBA test runner not found: $RUNNER_FILE"
  exit 1
fi

if [ ! -d "tests/spec/vba" ]; then
  echo "❌ tests/spec/vba directory not found"
  exit 1
fi

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
