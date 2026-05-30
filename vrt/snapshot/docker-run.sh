#!/bin/bash
set -euo pipefail

cd /workspace

LOCK_HASH=$(sha256sum pnpm-lock.yaml | cut -d' ' -f1)
HASH_FILE=node_modules/.pnpm-lock-hash

if [ -d node_modules ] && [ -f "$HASH_FILE" ] && [ "$(cat "$HASH_FILE")" = "$LOCK_HASH" ]; then
  echo "node_modules is up-to-date, skipping pnpm install"
else
  pnpm install --frozen-lockfile
  echo "$LOCK_HASH" > "$HASH_FILE"
fi
exec "$@"
