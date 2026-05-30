#!/bin/bash

# tests/spec/ 配下のすべてのテストを一括実行するスクリプト

echo "--- Starting all spec tests ---"

# テスト失敗フラグ
TESTS_FAILED=0

for f in tests/spec/*.test.ts; do
  echo "Running $f..."

  if ! npx tsx "$f"; then
    echo "❌ Test failed: $f"
    TESTS_FAILED=1
  fi
done

echo ""
if [ $TESTS_FAILED -eq 0 ]; then
  echo "--- All spec tests completed successfully! ---"
  exit 0
else
  echo "--- Some spec tests FAILED ---"
  exit 1
fi
