#!/bin/bash

# プロジェクト内のすべてのテストを一括実行するスクリプト
#  - tests/engine/           : コンパイラエンジン内部のテスト（Lexer, Parser, AST構造）
#  - tests/lsp/              : LSP（Language Server Protocol）のテスト
#  - tests/spec/             : VBA コンパイラの仕様テスト
#  - tests/test-libs-tests/  : test-libs/ (VBATest 等) のテスト
#  - sample/tests/ts/        : サンプル VBA コードのテスト
#
# Usage:
#   ./run_all_tests.sh      # 詳細出力を抑制（進捗表示のみ）
#   ./run_all_tests.sh -v   # 詳細出力（今までと同じ）

VERBOSE=0
if [ "$1" = "-v" ] || [ "$1" = "--verbose" ]; then
  VERBOSE=1
fi

if [ $VERBOSE -eq 0 ]; then
  echo -n "Running tests "
else
  echo "--- Starting all tests ---"
fi

# テスト失敗フラグ
TESTS_FAILED=0
FAILED_TESTS=""

for f in tests/engine/*.test.ts tests/lsp/*.test.ts tests/spec/*.test.ts tests/test-libs-tests/*.test.ts sample/tests/ts/*.test.ts; do
  # VBA テスト用の特別なランナーはスキップ（run_vba_tests.sh で実行）
  if [[ "$f" == *"run-all-vba-tests"* ]]; then
    continue
  fi
  if [ -f "$f" ]; then
    out="${f%.ts}.cjs"

    if [ $VERBOSE -eq 1 ]; then
      echo "Running $f..."
    else
      echo -n "."
    fi

    # esbuild でバンドル
    BUILD_OUTPUT=$(./node_modules/.bin/esbuild "$f" --bundle --outfile="$out" --platform=node 2>&1)
    BUILD_STATUS=$?

    if [ $BUILD_STATUS -ne 0 ]; then
      if [ $VERBOSE -eq 1 ]; then
        echo "❌ Build failed: $f"
      fi
      TESTS_FAILED=1
      FAILED_TESTS="${FAILED_TESTS}❌ Build failed: $f\n"
      continue
    fi

    # テスト実行（30秒タイムアウト）
    TEST_OUTPUT=$(timeout 30 node "$out" 2>&1)
    TEST_STATUS=$?

    if [ $TEST_STATUS -eq 124 ]; then
      if [ $VERBOSE -eq 1 ]; then
        echo "❌ Test timed out: $f"
      fi
      TESTS_FAILED=1
      FAILED_TESTS="${FAILED_TESTS}❌ Test timed out (>30s): $f\n"
      continue
    fi

    if [ $TEST_STATUS -ne 0 ]; then
      if [ $VERBOSE -eq 1 ]; then
        echo "❌ Test failed: $f"
        echo "$TEST_OUTPUT"
      fi
      TESTS_FAILED=1
      FAILED_TESTS="${FAILED_TESTS}❌ Test failed: $f\n${TEST_OUTPUT}\n"
    elif [ $VERBOSE -eq 1 ]; then
      echo "$TEST_OUTPUT"
    fi

    # 中間ファイルを削除したい場合はコメント解除
    # rm "$out"
  fi
done

echo ""
if [ $TESTS_FAILED -eq 0 ]; then
  echo "--- All tests completed successfully! ---"
  exit 0
else
  if [ $VERBOSE -eq 0 ]; then
    echo "--- Some tests FAILED ---"
    echo ""
    echo -e "$FAILED_TESTS"
  else
    echo "--- Some tests FAILED ---"
  fi
  exit 1
fi
