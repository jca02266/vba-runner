#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CJS="$SCRIPT_DIR/test-libs/vba-run.cjs"
TS="$SCRIPT_DIR/test-libs/vba-run.ts"
ESBUILD="$SCRIPT_DIR/node_modules/.bin/esbuild"

if [ ! -f "$CJS" ] || [ "$TS" -nt "$CJS" ]; then
    echo "Building vba-run..." >&2
    "$ESBUILD" "$TS" --bundle --outfile="$CJS" --platform=node >&2
fi

exec node "$CJS" "$@"
