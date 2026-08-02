#!/usr/bin/env bash
set -euo pipefail

root="${VBA_RUNNER_EVAL_WAIT_DIR:-${TMPDIR:-/tmp}/vba-runner-eval-wait}"
mkdir -p "$root"

usage() {
  printf 'usage: %s start <seconds> [id]\n' "$0"
  printf '       %s resume <id>\n' "$0"
  printf '       %s status [id]\n' "$0"
  printf '       %s stop <id>\n' "$0"
}

format_epoch() {
  date -r "$1" '+%Y-%m-%dT%H:%M:%S%z'
}

mark_done() {
  local dir=$1
  format_epoch "$(date +%s)" > "$dir/done"
  printf 'done\n' > "$dir/state"
  rm -f "$dir/pid"
}

launch_worker() {
  local dir=$1 deadline=$2
  nohup bash -c '
    dir=$1
    deadline=$2
    sleep_pid=
    on_interrupt() {
      [[ -n "$sleep_pid" ]] && kill "$sleep_pid" 2>/dev/null || true
      printf "interrupted\n" > "$dir/state"
      exit 143
    }
    trap on_interrupt TERM INT
    remaining=$((deadline - $(date +%s)))
    if (( remaining > 0 )); then
      sleep "$remaining" &
      sleep_pid=$!
      printf "%s\n" "$sleep_pid" > "$dir/pid"
      wait "$sleep_pid" || exit 143
    fi
    date -r "$(date +%s)" "+%Y-%m-%dT%H:%M:%S%z" > "$dir/done"
    printf "done\n" > "$dir/state"
    rm -f "$dir/pid"
  ' _ "$dir" "$deadline" > "$dir/output" 2>&1 &
}

read_state() {
  local id=$1 dir="$root/$1"
  if [[ ! -d "$dir" ]]; then
    printf 'unknown: %s\n' "$id" >&2
    return 1
  fi
  local state=pending pid='' deadline=''
  [[ -f "$dir/state" ]] && state=$(<"$dir/state")
  [[ -f "$dir/pid" ]] && pid=$(<"$dir/pid")
  [[ -f "$dir/resume_epoch" ]] && deadline=$(<"$dir/resume_epoch")
  if [[ "$state" == pending && -n "$deadline" && -n "$pid" ]] && ! kill -0 "$pid" 2>/dev/null; then
    if (( $(date +%s) >= deadline )); then
      mark_done "$dir"
      state=done
      pid=''
    else
      printf 'interrupted\n' > "$dir/state"
      state=interrupted
    fi
  elif [[ "$state" == interrupted && -n "$deadline" ]] && (( $(date +%s) >= deadline )); then
    mark_done "$dir"
    state=done
  fi
  printf '%s state=%s' "$id" "$state"
  [[ -n "$pid" ]] && printf ' pid=%s' "$pid"
  [[ -n "$deadline" ]] && printf ' resume_at=%s' "$(format_epoch "$deadline")"
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
    if [[ -e "$dir/state" ]]; then
      printf 'wait id already exists; use resume: %s\n' "$id" >&2
      exit 1
    fi
    mkdir -p "$dir"
    now=$(date +%s)
    deadline=$((now + seconds))
    format_epoch "$now" > "$dir/started_at"
    printf '%s\n' "$deadline" > "$dir/resume_epoch"
    printf 'pending\n' > "$dir/state"
    launch_worker "$dir" "$deadline"
    printf 'started id=%s resume_at=%s\n' "$id" "$(format_epoch "$deadline")"
    ;;
  resume)
    id=${2:-}
    [[ -n "$id" ]] || { usage >&2; exit 2; }
    dir="$root/$id"
    [[ -d "$dir" ]] || { printf 'unknown: %s\n' "$id" >&2; exit 1; }
    [[ -f "$dir/resume_epoch" ]] || { printf 'resume time unavailable: %s\n' "$id" >&2; exit 1; }
    if [[ -f "$dir/pid" ]] && kill -0 "$(<"$dir/pid")" 2>/dev/null; then
      printf 'wait already running: %s\n' "$id" >&2
      exit 1
    fi
    deadline=$(<"$dir/resume_epoch")
    if (( $(date +%s) >= deadline )); then
      mark_done "$dir"
      printf 'resumed id=%s state=done resume_at=%s\n' "$id" "$(format_epoch "$deadline")"
    else
      printf 'pending\n' > "$dir/state"
      launch_worker "$dir" "$deadline"
      printf 'resumed id=%s resume_at=%s\n' "$id" "$(format_epoch "$deadline")"
    fi
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
    if [[ -f "$dir/resume_epoch" ]]; then
      printf 'stopped id=%s resume_at=%s\n' "$id" "$(format_epoch "$(<"$dir/resume_epoch")")"
    else
      printf 'stopped id=%s\n' "$id"
    fi
    ;;
  *)
    usage >&2
    exit 2
    ;;
esac
