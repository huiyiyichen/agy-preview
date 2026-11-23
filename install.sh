#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

find_node() {
  for bin in "$(command -v node 2>/dev/null || true)" \
    "/Applications/Antigravity.app/Contents/Resources/app/bin/agy-node" \
    "$HOME/.local/share/antigravity/resources/app/bin/agy-node"; do
    if [ -n "$bin" ] && [ -x "$bin" ]; then
      if "$bin" -e 'const [a,b]=process.versions.node.split(".").map(Number); process.exit(a>22||(a===22&&b>=13)?0:1)'; then
        echo "$bin"
        return 0
      fi
    fi
  done
  echo "Error: Node.js >= 22.13.0 is required." >&2
  exit 1
}

NODE_BIN="$(find_node)"
exec "$NODE_BIN" "$REPO_DIR/scripts/install.mjs" "$@"
