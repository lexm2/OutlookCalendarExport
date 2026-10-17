#!/usr/bin/env bash
# Build the minified console script next to the readable source.
# Needs esbuild (pacman -S esbuild).
set -euo pipefail
cd "$(dirname "$0")"

src=outlook_calendar_export_console.js
out=outlook_calendar_export_console.min.js

esbuild "$src" \
  --minify \
  --target=esnext \
  --legal-comments=none \
  --log-level=warning \
  --outfile="$out"
truncate -s -1 "$out" # drop the trailing newline

echo "$out: $(wc -c < "$out") bytes (source $(wc -c < "$src") bytes)"
