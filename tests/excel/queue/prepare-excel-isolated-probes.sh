#!/usr/bin/env bash
set -euo pipefail

queue_dir=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
probe_dir="$queue_dir/isolated"
workbook_dir="$probe_dir/workbooks"
repo_dir=$(cd "$queue_dir/../../.." && pwd)
mkdir -p "$workbook_dir"
manifest_path="$workbook_dir/manifest.json"
source_dir=$(mktemp -d "${TMPDIR:-/tmp}/vba-runner-isolated-source.XXXXXX")
trap 'rm -rf "$source_dir"' EXIT

node - "$probe_dir" "$manifest_path" <<'NODE'
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const [dir, out] = process.argv.slice(2);
const names = fs.readdirSync(dir)
  .filter((name) => /^EV-\d{5}-.+\.bas$/i.test(name) || /^CONTROL\.bas$/i.test(name))
  .sort();
const normalize = (text) => text.replace(/\r\n/g, '\n');
const digest = (text) => crypto.createHash('sha256').update(text, 'utf8').digest('hex');
const entries = names.map((source) => {
  const text = normalize(fs.readFileSync(path.join(dir, source), 'utf8'));
  const module = text.match(/^Attribute VB_Name\s*=\s*"([^"]+)"/mi)?.[1];
  if (!module) throw new Error(`${source}: missing Attribute VB_Name`);
  return {
    id: source === 'CONTROL.bas' ? 'CONTROL' : text.match(/XL-\d{3}(?:-[A-Z0-9]+)*/i)?.[0],
    caseName: source === 'CONTROL.bas' ? 'CONTROL' : source.replace(/\.bas$/i, '').replace(/^EV-\d{5}-/, ''),
    source, module, procedure: 'RunCompileProbe',
    workbook: `workbooks/${source.replace(/\.bas$/i, '')}.xlsm`,
    sourceHash: digest(`${source}\n${text}\n`),
  };
});
const manifestText = entries.map((entry) => {
  const text = normalize(fs.readFileSync(path.join(dir, entry.source), 'utf8'));
  return `${entry.source}\n${text}\n`;
}).join('');
fs.writeFileSync(out, `${JSON.stringify({ version: 1, sourceHash: digest(manifestText), cases: entries }, null, 2)}\n`);
NODE

cp "$manifest_path" "$workbook_dir/manifest.next.json"
source_hash=$(node -e 'process.stdout.write(require(process.argv[1]).sourceHash)' "$workbook_dir/manifest.next.json")
if [[ -f "$manifest_path" ]] && cmp -s "$manifest_path" "$workbook_dir/manifest.next.json" \
  && [[ -f "$workbook_dir/source.sha256" ]] \
  && [[ "$(tr -d '\\r\\n' < "$workbook_dir/source.sha256")" == "$source_hash" ]] \
  && find "$workbook_dir" -maxdepth 1 -name '*.xlsm' -print -quit | grep -q .; then
  printf 'Isolated compile probes are current: %s\n' "$source_hash"
  rm -f "$workbook_dir/manifest.next.json"
  exit 0
fi
mv "$workbook_dir/manifest.next.json" "$manifest_path"
printf '%s\n' "$source_hash" > "$workbook_dir/source.sha256"

while IFS=$'\t' read -r source workbook; do
  rm -f "$probe_dir/$workbook"
  cp "$queue_dir/empty_with_macro.xlsm" "$probe_dir/$workbook"
  rm -rf "$source_dir"/*
  cp "$probe_dir/$source" "$source_dir/$source"
  (
    cd "$repo_dir"
    npm run vba-extractor -- import "$probe_dir/$workbook" "$source_dir" --yes
  )
done < <(node - "$manifest_path" <<'NODE'
const manifest = require(process.argv[2]);
for (const entry of manifest.cases) console.log(`${entry.source}\t${entry.workbook}`);
NODE
)

printf 'Prepared isolated compile probes: %s cases\nSource hash: %s\n' \
  "$(node -e 'process.stdout.write(String(require(process.argv[1]).cases.length))' "$manifest_path")" "$source_hash"
