#!/usr/bin/env bash
set -euo pipefail

queue_dir=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
stamp_file="$queue_dir/t.xlsm.source.sha256"
if [[ ! -f "$stamp_file" ]]; then
  echo "Missing preparation stamp: $stamp_file" >&2
  exit 1
fi

expected=$(tr -d '[:space:]' < "$stamp_file")
actual=$(node - "$queue_dir" <<'NODE'
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const directory = process.argv[2];
const names = fs.readdirSync(directory)
  .filter((name) => /\.(?:bas|cls|frm)$/i.test(name))
  .sort();
const hashScheme = 'QUEUE_SOURCE_HASH_SCHEME=grouped-v1\n';
const normalize = (text) => text.replace(/\r\n/g, '\n')
  .replace(/^\s*Private Const QUEUE_SOURCE_SHA256\s+As String\s*=\s*"[^"]*"\s*$/mi,
    'Private Const QUEUE_SOURCE_SHA256 As String = "__QUEUE_SOURCE_SHA256__"');
const manifest = hashScheme + names.map((name) => `${name}\n${normalize(
  fs.readFileSync(path.join(directory, name), 'utf8'))}\n`).join('');
process.stdout.write(crypto.createHash('sha256').update(manifest, 'utf8').digest('hex'));
NODE
)

if [[ ! "$expected" =~ ^[0-9a-fA-F]{64}$ || "$expected" != "$actual" ]]; then
  echo "Stale Excel queue workbook. Run prepare-excel-vba.sh first." >&2
  echo "Prepared=$expected Current=$actual" >&2
  exit 1
fi
printf 'Excel queue source matches prepared workbook: %s\n' "$actual"
