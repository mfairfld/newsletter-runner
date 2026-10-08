#!/bin/sh
set -u

export PATH=/usr/bin:/bin:/usr/sbin:/sbin
export LANG=en_US.UTF-8

REPO="/Users/Workspace/code/newsletter"
PY="$REPO/.venv/bin/python"
LOGDIR="/Users/Workspace/Library/Logs/newsletter"
LOG="$LOGDIR/newsletter.log"
MAXBYTES=1048576

/bin/mkdir -p "$LOGDIR"

if [ -f "$LOG" ]; then
  SIZE=$(/usr/bin/stat -f%z "$LOG")
  if [ "$SIZE" -gt "$MAXBYTES" ]; then
    /bin/mv -f "$LOG" "$LOG.1"
  fi
fi

exec >> "$LOG" 2>&1

echo "=== start $(/bin/date '+%Y-%m-%d %H:%M:%S %Z') ==="

if [ ! -x "$PY" ]; then
  echo "FATAL: interpreter missing at $PY"
  echo "=== end status=127 $(/bin/date '+%Y-%m-%d %H:%M:%S %Z') ==="
  exit 127
fi

cd "$REPO" || exit 1
"$PY" newsletter_bot.py
STATUS=$?

echo "=== end status=$STATUS $(/bin/date '+%Y-%m-%d %H:%M:%S %Z') ==="
exit $STATUS
