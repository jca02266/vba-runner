#!/usr/bin/env bash
set -euo pipefail

root="${VBA_RUNNER_EVAL_WAIT_DIR:-${TMPDIR:-/tmp}/vba-runner-eval-wait}"
mkdir -p "$root"

usage() {
  printf 'usage: %s start <seconds> [id]\n' "$0"
  printf '       %s status [id]\n' "$0"
  printf '       %s stop <id>\n' "$0"
}

read_state() {
  local dir="$root/$1"
  if [[ ! -d "$dir" ]]; then
    printf 'unknown: %s\n' "$1" >&2
    return 1
  fi
  local state=pending pid=''
  [[ -f "$dir/state" ]] && state=$(<"$dir/state")
  [[ -f "$dir/pid" ]] && pid=$(<"$dir/pid")
  if [[ "$state" == pending && -n "$pid" ]] && ! kill -0 "$pid" 2>/dev/null; then
    state=interrupted
  fi
  printf '%s state=%s' "$1" "$state"
  [[ -n "$pid" ]] && printf ' pid=%s' "$pid"
  [[ -f "$dir/done" ]] && printf ' done=%s' "$(<"$dir/done")"
  printf '\n'
}

command=${1:-}
case "$command" in
  start)
    seconds=${2:-1800}
    [[ "$seconds" =~ ^[1-9][0-9]*$ ]] || { printf 'seconds must be a positive integer\n' >&2; exit 2; }
    id=${3:-"eval-$(date +%Y%m%d-%H%M%S)-$$"}
    dir="$root/$id"
    if [[ -e "$dir/state" ]] && [[ "$(<"$dir/state")" == pending ]] && [[ -f "$dir/pid" ]] && kill -0 "$(<"$dir/pid")" 2>/dev/null; then
      printf 'wait already running: %s\n' "$id" >&2
      exit 1
    fi
    mkdir -p "$dir"
    printf 'pending\n' > "$dir/state"
    nohup bash -c '
      dir=$1
      seconds=$2
      on_interrupt() { printf "interrupted\n" > "$dir/state"; exit 143; }
      trap on_interrupt TERM INT
      sleep "$seconds"
      date "+%Y-%m-%dT%H:%M:%S%z" > "$dir/done"
      printf "done\n" > "$dir/state"
      rm -f "$dir/pid"
    ' _ "$dir" "$seconds" > "$dir/output" 2>&1 &
    printf '%s\n' "$!" > "$dir/pid"
    printf 'started id=%s seconds=%s\n' "$id" "$seconds"
    ;;
  status)
    if [[ -n "${2:-}" ]]; then
      read_state "$2"
    else
      found=0
      for dir in "$root"/*; do
        [[ -d "$dir" ]] || continue
        found=1
        read_state "${dir##*/}"
      done
      (( found )) || printf 'no waits\n'
    fi
    ;;
  stop)
    id=${2:-}
    [[ -n "$id" ]] || { usage >&2; exit 2; }
    dir="$root/$id"
    [[ -d "$dir" ]] || { printf 'unknown: %s\n' "$id" >&2; exit 1; }
    if [[ -f "$dir/pid" ]]; then
      kill "$(<"$dir/pid")" 2>/dev/null || true
    fi
    printf 'interrupted\n' > "$dir/state"
    printf 'stopped id=%s\n' "$id"
    ;;
  *)
    usage >&2
    exit 2
    ;;
esac
