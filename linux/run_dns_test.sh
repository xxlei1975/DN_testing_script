#!/usr/bin/env bash
# ============================================================
#  Africa_DN_Testing - one-click launcher (Linux)
#
#  1. reuses the interpreter recorded in cache/python_path.txt, but ONLY when
#     that record was written on this very machine and the interpreter still
#     works. A cache folder copied from another PC is ignored instead of being
#     trusted (it contains that PC's absolute paths).
#  2. reuses dig from PATH or from tools/bind/
#  3. if anything is missing, calls scripts/setup_env.sh which installs it
#     automatically (Python 3 / venv / dependencies / dig via the distro
#     package manager)
#  4. finally runs domain_dns_test.py with all given arguments
#
#  All text in this file is ASCII on purpose: it stays readable no matter what
#  LANG/LC_ALL is set to (a non-UTF-8 locale would turn Chinese literals into
#  mojibake). Chinese messages come from the Python program instead.
#
#  Examples
#      ./run_dns_test.sh                  full run
#      ./run_dns_test.sh --limit 5        first 5 domains only
#      ./run_dns_test.sh --dry-run        preview, no file written
#      ./run_dns_test.sh --inplace        write back to the original file
#
#  Exit codes: same as domain_dns_test.py (0 ok, 2 input/dig problem,
#  3 cannot save the workbook); 1 means the environment could not be prepared.
# ============================================================
set -u

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PYFILE="$SCRIPT_DIR/cache/python_path.txt"
SETUP="$SCRIPT_DIR/scripts/setup_env.sh"
PROGRAM="$SCRIPT_DIR/domain_dns_test.py"

PY=""
FALLBACK_PY=""

log()  { printf '[INFO] %s\n' "$*"; }
warn() { printf '[WARN] %s\n' "$*"; }
err()  { printf '[ERROR] %s\n' "$*" >&2; }

cd "$SCRIPT_DIR" || { err "cannot enter $SCRIPT_DIR"; exit 1; }

# ---------- helpers ----------

# an interpreter is only accepted when it exists AND can import the deps
interpreter_works() {
    local py="$1"
    [ -n "$py" ] || return 1
    [ -x "$py" ] || return 1
    "$py" -c 'import openpyxl, requests' >/dev/null 2>&1
}

# reads cache/python_path.txt; the host name must match this machine
load_cached_python() {
    PY=""
    [ -f "$PYFILE" ] || return 0

    local machine="" python="" key value
    while IFS='=' read -r key value; do
        case "$key" in
            machine) machine="$value" ;;
            python)  python="$value" ;;
        esac
    done < "$PYFILE"

    [ -n "$python" ] || return 0

    local host
    host="$(uname -n 2>/dev/null || printf '')"
    if [ -n "$machine" ] && [ -n "$host" ] && [ "$machine" != "$host" ]; then
        log "cache/python_path.txt was written on host $machine - ignoring it here."
        return 0
    fi
    if ! interpreter_works "$python"; then
        log "the recorded interpreter is not usable here - ignoring it."
        return 0
    fi
    PY="$python"
}

# last resort: python3 found on PATH (the dependencies may be missing, which
# is fine here: setup_env.sh installs them into the project venv afterwards)
find_python_on_path() {
    local name
    for name in python3 python3.13 python3.12 python3.11 python3.10 python3.9 python; do
        if command -v "$name" >/dev/null 2>&1; then
            command -v "$name"
            return 0
        fi
    done
    printf ''
}

dig_available() {
    command -v dig >/dev/null 2>&1 && return 0
    [ -x "$SCRIPT_DIR/tools/bind/dig" ] && return 0
    return 1
}

# ---------- step 1: interpreter recorded by a previous run ----------
load_cached_python
if [ -n "$PY" ] && dig_available; then
    log "using the cached interpreter"
else
    # ---------- step 2: provision this machine ----------
    log "checking the environment - Python / dependencies / dig may be installed now"
    printf '\n'
    if [ -x "$SETUP" ]; then
        bash "$SETUP" || warn "setup_env.sh reported a problem - trying to continue anyway"
    else
        warn "scripts/setup_env.sh is missing - skipping automatic provisioning"
    fi
    printf '\n'

    load_cached_python
    if [ -z "$PY" ]; then
        FALLBACK_PY="$(find_python_on_path)"
        if [ -n "$FALLBACK_PY" ]; then
            warn "falling back to $FALLBACK_PY (dependencies may be missing)"
            PY="$FALLBACK_PY"
        fi
    fi
fi

if [ -z "$PY" ]; then
    err "Python is still unavailable. Run this by hand to see the details:"
    err "    bash scripts/setup_env.sh"
    exit 1
fi

# ---------- step 3: run ----------
log "Python  : $PY"
log "Program : $PROGRAM"
printf '\n'

"$PY" "$PROGRAM" "$@"
rc=$?

printf '\n'
if [ "$rc" -eq 0 ]; then
    printf '[DONE] Finished successfully.\n'
else
    printf '[FAILED] exit code %s - see the messages above.\n' "$rc"
fi
exit "$rc"
