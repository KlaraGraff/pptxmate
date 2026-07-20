#!/usr/bin/env bash

set -euo pipefail

[[ -n "${HOME:-}" ]] || {
  printf 'PPTXMate watcher installer: HOME is not set\n' >&2
  exit 1
}

LABEL="com.local.pptxmate-powerpoint-dev-server"
OLD_LABEL="com.local.openppt-powerpoint-dev-server"
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
REPO_ROOT="$(cd -- "$SCRIPT_DIR/.." && pwd -P)"
WATCHER="$SCRIPT_DIR/watch-powerpoint-dev-server.mjs"
LAUNCH_AGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST="$LAUNCH_AGENTS_DIR/$LABEL.plist"
OLD_PLIST="$LAUNCH_AGENTS_DIR/$OLD_LABEL.plist"
DOMAIN="gui/$UID"

usage() {
  cat <<'EOF'
Install the optional PPTXMate PowerPoint lifecycle watcher on macOS.

Usage:
  ./scripts/install-macos-powerpoint-watcher.sh [options]

Options:
  --node PATH          Node executable (default: PPTXMATE_NODE_BIN or PATH lookup)
  --pnpm PATH          pnpm executable (default: PPTXMATE_PNPM_BIN or PATH lookup)
  --path VALUE         PATH for the watcher/dev server (default: PPTXMATE_PATH or PATH)
  --log-dir PATH       Log directory (default: ~/Library/Logs/PPTXMate)
  --interval-ms VALUE  PowerPoint polling interval (default: 3000)
  --port VALUE         Dev-server port guard (default: 3001)
  --cc-switch-url URL  CC Switch origin (default: http://127.0.0.1:15721)
  --no-cc-switch       Disable the local CC Switch /v1 proxy
  --dry-run            Print the generated LaunchAgent without installing it
  -h, --help           Show this help

Environment equivalents:
  PPTXMATE_NODE_BIN, PPTXMATE_PNPM_BIN, PPTXMATE_PATH, PPTXMATE_LOG_DIR,
  PPTXMATE_CHECK_INTERVAL_MS, PPTXMATE_STOP_TIMEOUT_MS, PPTXMATE_PORT,
  PPTXMATE_CC_SWITCH_URL, PPTXMATE_CC_SWITCH_ENABLED

Installing replaces the previous PPTXMate agent and removes the legacy
com.local.openppt-powerpoint-dev-server LaunchAgent if it is present.
EOF
}

die() {
  printf 'PPTXMate watcher installer: %s\n' "$*" >&2
  exit 1
}

require_value() {
  [[ $# -ge 2 && -n "$2" ]] || die "$1 requires a value"
}

require_integer() {
  local name="$1"
  local value="$2"
  local minimum="$3"
  local maximum="$4"
  [[ "$value" =~ ^[0-9]+$ ]] || die "$name must be an integer"
  ((value >= minimum && value <= maximum)) || die "$name must be between $minimum and $maximum"
}

xml_escape() {
  local value="$1"
  [[ "$value" != *$'\n'* && "$value" != *$'\r'* ]] || die "paths and environment values cannot contain newlines"
  printf '%s' "$value" | /usr/bin/sed \
    -e 's/&/\&amp;/g' \
    -e 's/</\&lt;/g' \
    -e 's/>/\&gt;/g' \
    -e 's/"/\&quot;/g' \
    -e "s/'/\&apos;/g"
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

RUNTIME_PATH="${PPTXMATE_PATH:-${PATH:-/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin}}"
NODE_CANDIDATE="${PPTXMATE_NODE_BIN:-node}"
PNPM_CANDIDATE="${PPTXMATE_PNPM_BIN:-pnpm}"
LOG_DIR="${PPTXMATE_LOG_DIR:-$HOME/Library/Logs/PPTXMate}"
CHECK_INTERVAL_MS="${PPTXMATE_CHECK_INTERVAL_MS:-3000}"
STOP_TIMEOUT_MS="${PPTXMATE_STOP_TIMEOUT_MS:-5000}"
PORT="${PPTXMATE_PORT:-3001}"
CC_SWITCH_URL="${PPTXMATE_CC_SWITCH_URL:-http://127.0.0.1:15721}"
CC_SWITCH_ENABLED="${PPTXMATE_CC_SWITCH_ENABLED:-1}"
DRY_RUN=false

while (($#)); do
  case "$1" in
    --node)
      require_value "$@"
      NODE_CANDIDATE="$2"
      shift 2
      ;;
    --pnpm)
      require_value "$@"
      PNPM_CANDIDATE="$2"
      shift 2
      ;;
    --path)
      require_value "$@"
      RUNTIME_PATH="$2"
      shift 2
      ;;
    --log-dir)
      require_value "$@"
      LOG_DIR="$2"
      shift 2
      ;;
    --interval-ms)
      require_value "$@"
      CHECK_INTERVAL_MS="$2"
      shift 2
      ;;
    --port)
      require_value "$@"
      PORT="$2"
      shift 2
      ;;
    --cc-switch-url)
      require_value "$@"
      CC_SWITCH_URL="$2"
      CC_SWITCH_ENABLED=1
      shift 2
      ;;
    --no-cc-switch)
      CC_SWITCH_ENABLED=0
      shift
      ;;
    --dry-run)
      DRY_RUN=true
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
((EUID != 0)) || die "run this installer as the signed-in user, not with sudo"
[[ -f "$WATCHER" ]] || die "watcher not found: $WATCHER"
[[ -f "$REPO_ROOT/package.json" ]] || die "package.json not found in repository root: $REPO_ROOT"
[[ -n "$RUNTIME_PATH" ]] || die "PATH cannot be empty"
require_safe_log_dir "$LOG_DIR"
require_integer "poll interval" "$CHECK_INTERVAL_MS" 250 3600000
require_integer "stop timeout" "$STOP_TIMEOUT_MS" 250 120000
require_integer "port" "$PORT" 1 65535

resolve_executable() {
  local candidate="$1"
  local resolved
  if [[ "$candidate" == */* ]]; then
    [[ -x "$candidate" ]] || die "executable not found or not executable: $candidate"
    resolved="$(cd -- "$(dirname -- "$candidate")" && pwd -P)/$(basename -- "$candidate")"
  else
    resolved="$(PATH="$RUNTIME_PATH" command -v "$candidate" 2>/dev/null || true)"
    [[ -n "$resolved" && -x "$resolved" ]] || die "$candidate was not found in the configured PATH"
  fi
  printf '%s' "$resolved"
}

NODE_BIN="$(resolve_executable "$NODE_CANDIDATE")"
PNPM_BIN="$(resolve_executable "$PNPM_CANDIDATE")"
WATCHER_LOG="$LOG_DIR/powerpoint-watcher.log"
LAUNCH_AGENT_LOG="$LOG_DIR/launch-agent.log"
LOG_MARKER="$LOG_DIR/.pptxmate-watcher-logs"

TMP_PLIST="$(/usr/bin/mktemp "${TMPDIR:-/tmp}/pptxmate-powerpoint-watcher.XXXXXX")"
cleanup() {
  rm -f "$TMP_PLIST"
}
trap cleanup EXIT

ESCAPED_LABEL="$(xml_escape "$LABEL")"
ESCAPED_NODE="$(xml_escape "$NODE_BIN")"
ESCAPED_WATCHER="$(xml_escape "$WATCHER")"
ESCAPED_REPO="$(xml_escape "$REPO_ROOT")"
ESCAPED_PNPM="$(xml_escape "$PNPM_BIN")"
ESCAPED_PATH="$(xml_escape "$RUNTIME_PATH")"
ESCAPED_WATCHER_LOG="$(xml_escape "$WATCHER_LOG")"
ESCAPED_AGENT_LOG="$(xml_escape "$LAUNCH_AGENT_LOG")"
ESCAPED_CC_SWITCH_URL="$(xml_escape "$CC_SWITCH_URL")"
ESCAPED_CC_SWITCH_ENABLED="$(xml_escape "$CC_SWITCH_ENABLED")"

cat >"$TMP_PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>$ESCAPED_LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$ESCAPED_NODE</string>
    <string>$ESCAPED_WATCHER</string>
  </array>
  <key>WorkingDirectory</key>
  <string>$ESCAPED_REPO</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>PPTXMATE_NODE_BIN</key>
    <string>$ESCAPED_NODE</string>
    <key>PPTXMATE_PNPM_BIN</key>
    <string>$ESCAPED_PNPM</string>
    <key>PPTXMATE_PATH</key>
    <string>$ESCAPED_PATH</string>
    <key>PPTXMATE_WATCHER_LOG</key>
    <string>$ESCAPED_WATCHER_LOG</string>
    <key>PPTXMATE_CHECK_INTERVAL_MS</key>
    <string>$CHECK_INTERVAL_MS</string>
    <key>PPTXMATE_STOP_TIMEOUT_MS</key>
    <string>$STOP_TIMEOUT_MS</string>
    <key>PPTXMATE_PORT</key>
    <string>$PORT</string>
    <key>PPTXMATE_CC_SWITCH_URL</key>
    <string>$ESCAPED_CC_SWITCH_URL</string>
    <key>PPTXMATE_CC_SWITCH_ENABLED</key>
    <string>$ESCAPED_CC_SWITCH_ENABLED</string>
  </dict>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>ProcessType</key>
  <string>Background</string>
  <key>StandardOutPath</key>
  <string>$ESCAPED_AGENT_LOG</string>
  <key>StandardErrorPath</key>
  <string>$ESCAPED_AGENT_LOG</string>
</dict>
</plist>
EOF

/usr/bin/plutil -lint "$TMP_PLIST" >/dev/null || die "generated LaunchAgent plist is invalid"

if [[ "$DRY_RUN" == true ]]; then
  cat "$TMP_PLIST"
  exit 0
fi

mkdir -p "$LAUNCH_AGENTS_DIR"
if [[ -e "$LOG_DIR" && ! -d "$LOG_DIR" ]]; then
  die "log directory path is not a directory: $LOG_DIR"
fi
if [[ -d "$LOG_DIR" && ! -f "$LOG_MARKER" ]] && [[ -n "$(find "$LOG_DIR" -mindepth 1 -maxdepth 1 -print -quit)" ]]; then
  die "refusing to use a non-empty directory not created for PPTXMate logs: $LOG_DIR"
fi
mkdir -p -m 700 "$LOG_DIR"
chmod 700 "$LOG_DIR"
printf 'PPTXMate PowerPoint watcher logs\n' >"$LOG_MARKER"
touch "$WATCHER_LOG" "$LAUNCH_AGENT_LOG"
chmod 600 "$LOG_MARKER" "$WATCHER_LOG" "$LAUNCH_AGENT_LOG"

bootout_agent() {
  local label="$1"
  local plist="$2"
  if /bin/launchctl print "$DOMAIN/$label" >/dev/null 2>&1; then
    /bin/launchctl bootout "$DOMAIN/$label" || die "could not stop LaunchAgent $label"
  elif [[ -f "$plist" ]]; then
    /bin/launchctl bootout "$DOMAIN" "$plist" >/dev/null 2>&1 || true
  fi
}

# Migrate the pre-PPTXMate label before loading the replacement agent.
bootout_agent "$OLD_LABEL" "$OLD_PLIST"
rm -f "$OLD_PLIST"
bootout_agent "$LABEL" "$PLIST"

/usr/bin/install -m 600 "$TMP_PLIST" "$PLIST"
/bin/launchctl enable "$DOMAIN/$LABEL"
if ! /bin/launchctl bootstrap "$DOMAIN" "$PLIST"; then
  die "LaunchAgent installation failed; inspect $LAUNCH_AGENT_LOG"
fi

printf 'Installed PPTXMate PowerPoint watcher.\n'
printf 'LaunchAgent: %s\n' "$PLIST"
printf 'Watcher log: %s\n' "$WATCHER_LOG"
printf 'Uninstall with: %s/scripts/uninstall-macos-powerpoint-watcher.sh\n' "$REPO_ROOT"
