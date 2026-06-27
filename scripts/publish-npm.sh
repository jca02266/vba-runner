#!/usr/bin/env bash
# npm publish wrapper — writes auth token to .npmrc instead of CLI args.
# Run from the npm package directory (CWD must contain .env and package.json).
set -euo pipefail

ENV_FILE=".env"
NPMRC=".npmrc"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Error: $ENV_FILE not found in $(pwd)" >&2
  exit 1
fi

NPM_TOKEN=$(grep -v '^#' "$ENV_FILE" | grep '^NPM_TOKEN=' | cut -d= -f2-)

if [[ -z "$NPM_TOKEN" ]]; then
  echo "Error: NPM_TOKEN not found in $ENV_FILE" >&2
  exit 1
fi

trap 'rm -f "$NPMRC"' EXIT INT TERM
echo "//registry.npmjs.org/:_authToken=${NPM_TOKEN}" > "$NPMRC"

npm publish --access public
