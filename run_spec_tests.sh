#!/bin/bash

# tests/spec/ 配下のすべてのテストを一括実行するスクリプト

# 失敗時に停止
set -e

echo "--- Starting all spec tests ---"

for f in tests/spec/*.test.ts; do
  out="${f%.ts}.cjs"
  echo "Running $f..."
  npx esbuild "$f" --bundle --outfile="$out" --platform=node
  node "$out"
  # 中間ファイルを削除したい場合はコメント解除
  # rm "$out"
done

echo "--- All spec tests completed successfully! ---"
