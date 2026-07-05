#!/usr/bin/env bash
# npm publish wrapper — writes auth token to .npmrc instead of CLI args.
# Usage: publish-npm.sh [path/to/.env]
#   Default .env location: .env in CWD (i.e. the npm package directory).
#   Pass an explicit path when the .env lives elsewhere (e.g. a sibling package).
set -euo pipefail

ENV_FILE="${1:-.env}"
NPMRC=".npmrc"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Error: $ENV_FILE not found" >&2
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
