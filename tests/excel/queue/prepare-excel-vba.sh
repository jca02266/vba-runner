#!/usr/bin/env bash
set -euo pipefail

queue_dir=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
repo_dir=$(cd "$queue_dir/../../.." && pwd)
workbook="$queue_dir/t.xlsm"

cp "$queue_dir/empty_with_macro.xlsm" "$workbook"
(
  cd "$repo_dir"
  npm run vba-extractor -- import "$workbook" "$queue_dir" --yes
)

printf 'Prepared workbook: %s\n' "$workbook"
