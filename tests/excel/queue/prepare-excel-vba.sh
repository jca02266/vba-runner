#!/usr/bin/env bash
set -euo pipefail

queue_dir=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
repo_dir=$(cd "$queue_dir/../../.." && pwd)
workbook="$queue_dir/t.xlsm"
workbook_stamp="$workbook.source.sha256"
stamped_dir=$(mktemp -d "${TMPDIR:-/tmp}/vba-runner-excel-source.XXXXXX")
trap 'rm -rf "$stamped_dir"' EXIT

source_hash=$(node - "$queue_dir" "$stamped_dir" <<'NODE'
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const [sourceDir, stampedDir] = process.argv.slice(2);
const names = fs.readdirSync(sourceDir)
  .filter((name) => /\.(?:bas|cls|frm)$/i.test(name))
  .sort();
if (names.length === 0) throw new Error(`No VBA source files found: ${sourceDir}`);
const sourcePrefix = (name) => {
  for (const prefix of ['ExcelQueue', 'FormatMatrix', 'RadixMatrix']) {
    if (name.toLowerCase().startsWith(prefix.toLowerCase())) return prefix;
  }
  return null;
};
const hashScheme = 'QUEUE_SOURCE_HASH_SCHEME=grouped-v1\n';
const normalize = (text) => text.replace(/\r\n/g, '\n')
  .replace(/^\s*Private Const QUEUE_SOURCE_SHA256\s+As String\s*=\s*"[^"]*"\s*$/mi,
    'Private Const QUEUE_SOURCE_SHA256 As String = "__QUEUE_SOURCE_SHA256__"');
const hashFor = (groupNames, includeScheme = false) => {
  const manifest = (includeScheme ? hashScheme : '') + groupNames.map((name) => `${name}\n${normalize(
    fs.readFileSync(path.join(sourceDir, name), 'utf8'))}\n`).join('');
  return crypto.createHash('sha256').update(manifest, 'utf8').digest('hex');
};
const groups = new Map();
for (const name of names) {
  const prefix = sourcePrefix(name);
  if (prefix) groups.set(prefix, [...(groups.get(prefix) ?? []), name]);
}
const groupHashes = new Map([...groups].map(([prefix, groupNames]) => [prefix, hashFor(groupNames)]));
const hash = hashFor(names, true);
for (const name of names) {
  let text = fs.readFileSync(path.join(sourceDir, name), 'utf8');
  if (/\.bas$/i.test(name)) {
    const prefix = sourcePrefix(name);
    const groupHash = prefix ? groupHashes.get(prefix) : null;
    const replaced = text.replace(/^\s*Private Const QUEUE_SOURCE_SHA256\s+As String\s*=\s*"[^"]*"\s*$/mi,
      `Private Const QUEUE_SOURCE_SHA256 As String = "${groupHash ?? hash}"`);
    if (replaced === text && [
      'excelqueueverification.bas',
      'formatmatrixverification.bas',
      'radixmatrixverification.bas',
    ].includes(name.toLowerCase())) {
      throw new Error('QUEUE_SOURCE_SHA256 marker not found');
    }
    text = replaced;
  }
  fs.writeFileSync(path.join(stampedDir, name), text);
}
process.stdout.write(hash);
NODE
)

cp "$queue_dir/empty_with_macro.xlsm" "$workbook"
(
  cd "$repo_dir"
  npm run vba-extractor -- import "$workbook" "$stamped_dir" --yes
)

printf 'Prepared workbook: %s\nSource hash: %s\n' "$workbook" "$source_hash"
printf '%s\n' "$source_hash" > "$workbook_stamp"
printf 'Preparation stamp: %s\n' "$workbook_stamp"

# Compile-error candidates use one workbook per case so a VBE compile failure
# cannot prevent the normal queue from running.  Keep their preparation under
# the same development-side command; Windows still receives prepared files.
"$queue_dir/prepare-excel-isolated-probes.sh"
