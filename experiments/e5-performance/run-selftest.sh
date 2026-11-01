#!/usr/bin/env bash
# Runs a Prime Word browser self-test in headless Edge against the dev server (npm run dev) and
# prints its "[selftest]" lines.
#   T=perf  F=<docx> ./run-selftest.sh   typing speed + page alignment (default: long.docx)
#   T=align F=<docx> ./run-selftest.sh   page alignment only
# The document must be inside the repository (the dev server serves only files there).
set -u
here="$(cd "$(dirname "$0")" && pwd -W 2>/dev/null || pwd)"
EDGE="${EDGE:-/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe}"
T="${T:-perf}"; F="${F:-$here/long.docx}"
tmp="${TMPDIR:-${TEMP:-/tmp}}/prime-word-selftest"; mkdir -p "$tmp"; log="$tmp/console.log"; rm -f "$log"
("$EDGE" --headless=new --disable-gpu --user-data-dir="$tmp/profile" --enable-logging=stderr --v=0 \
  --window-size=1400,1000 "http://localhost:5173/?selftest=$T&file=$F" > "$log" 2>&1 &)
for _ in $(seq 1 40); do grep -aq "selftest\] align" "$log" && break; sleep 3; done
grep -a "selftest\]\|Uncaught" "$log" | sed 's/.*\] "//; s/", source.*//'
# Stop only the test browser (identified by its profile folder), never the user's own Edge windows.
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { \$_.CommandLine -like '*prime-word-selftest*' } | ForEach-Object { Stop-Process -Id \$_.ProcessId -Force -ErrorAction SilentlyContinue }"
