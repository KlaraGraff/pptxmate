#!/usr/bin/env bash

set -euo pipefail

[[ -n "${HOME:-}" ]] || {
  printf 'PPTXMate watcher uninstaller: HOME is not set\n' >&2
  exit 1
}

LABEL="com.local.pptxmate-powerpoint-dev-server"
OLD_LABEL="com.local.openppt-powerpoint-dev-server"
DOMAIN="gui/$UID"
LAUNCH_AGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST="$LAUNCH_AGENTS_DIR/$LABEL.plist"
OLD_PLIST="$LAUNCH_AGENTS_DIR/$OLD_LABEL.plist"
LOG_DIR="${PPTXMATE_LOG_DIR:-$HOME/Library/Logs/PPTXMate}"
REMOVE_LOGS=false

usage() {
  cat <<'EOF'
Uninstall the optional PPTXMate PowerPoint lifecycle watcher on macOS.

Usage:
  ./scripts/uninstall-macos-powerpoint-watcher.sh [--remove-logs]

Options:
  --remove-logs  Remove PPTXMate's log files and its directory if then empty
  -h, --help     Show this help

The uninstaller stops only the watcher and the dev-server process group that
the watcher owns. It never searches for or kills arbitrary port users.
EOF
}

die() {
  printf 'PPTXMate watcher uninstaller: %s\n' "$*" >&2
  exit 1
}

require_safe_log_dir() {
  local value="$1"
  local normalized="$value"
  while [[ "$normalized" != "/" && "$normalized" == */ ]]; do
    normalized="${normalized%/}"
  done
  [[ -n "$value" ]] || die "log directory cannot be empty"
  [[ "$value" == /* ]] || die "log directory must be an absolute path"
  [[ "$normalized" != "/" && "$normalized" != "${HOME%/}" ]] || die "refusing unsafe log directory: $value"
  [[ "/$value/" != *"/../"* ]] || die "log directory cannot contain '..' path components"
  [[ "/$value/" != *"/./"* ]] || die "log directory cannot contain '.' path components"
  [[ "$value" != *$'\n'* && "$value" != *$'\r'* ]] || die "log directory cannot contain newlines"
}

while (($#)); do
  case "$1" in
    --remove-logs)
      REMOVE_LOGS=true
      shift
      ;;
    -h | --help)
      usage
      exit 0
      ;;
    *)
      die "unknown option: $1 (use --help)"
      ;;
  esac
done

if [[ "${PPTXMATE_LOG_DIR+x}" == "x" && -z "${PPTXMATE_LOG_DIR:-}" ]]; then
  die "PPTXMATE_LOG_DIR cannot be empty"
fi
[[ "$(uname -s)" == "Darwin" ]] || die "this optional watcher supports macOS only"
((EUID != 0)) || die "run this uninstaller as the signed-in user, not with sudo"

bootout_agent() {
  local label="$1"
  local plist="$2"
  if /bin/launchctl print "$DOMAIN/$label" >/dev/null 2>&1; then
    /bin/launchctl bootout "$DOMAIN/$label" || die "could not stop LaunchAgent $label"
  elif [[ -f "$plist" ]]; then
    /bin/launchctl bootout "$DOMAIN" "$plist" >/dev/null 2>&1 || true
  fi
}

bootout_agent "$LABEL" "$PLIST"
bootout_agent "$OLD_LABEL" "$OLD_PLIST"
rm -f "$PLIST" "$OLD_PLIST"

if [[ "$REMOVE_LOGS" == true ]]; then
  require_safe_log_dir "$LOG_DIR"
  LOG_MARKER="$LOG_DIR/.pptxmate-watcher-logs"
  [[ -f "$LOG_MARKER" ]] || die "refusing to remove unmarked log directory: $LOG_DIR"
  rm -f "$LOG_DIR/powerpoint-watcher.log" "$LOG_DIR/launch-agent.log" "$LOG_MARKER"
  rmdir "$LOG_DIR" 2>/dev/null ||
    printf 'Kept non-empty log directory: %s\n' "$LOG_DIR"
fi

printf 'Uninstalled PPTXMate PowerPoint watcher.\n'
if [[ "$REMOVE_LOGS" == false ]]; then
  printf 'Logs were kept at %s (use --remove-logs to delete them).\n' "$LOG_DIR"
fi
