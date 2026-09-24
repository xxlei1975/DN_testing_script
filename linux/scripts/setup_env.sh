#!/usr/bin/env bash
# ============================================================
#  Africa_DN_Testing - environment provisioning (Linux)
#
#  What this script does
#    1. finds a usable Python 3 (>= 3.9)
#    2. installs python3 / pip / venv with the distribution package
#       manager when they are missing (apt / dnf / yum / apk / zypper /
#       pacman are autodetected)
#    3. creates or repairs <project>/.venv
#    4. installs the dependencies (openpyxl, requests): local wheels/ first
#       (works offline), then PyPI
#    5. makes sure 'dig' is available (dnsutils / bind-utils / bind-tools)
#    6. checks whether the current user may send ICMP (otherwise ping needs
#       a one-time fix, see the hint printed at the end)
#    7. records the resolved interpreter in cache/python_path.txt together
#       with the host name, so a cache folder copied from another machine is
#       never trusted
#
#  System packages are only installed when the script can run as root or via
#  sudo; otherwise it prints the exact commands for you to run by hand.
#
#  Usage
#      scripts/setup_env.sh              provision this machine
#      scripts/setup_env.sh --no-install  only diagnose, never touch the system
#      scripts/setup_env.sh --help
#
#  All text in this file is ASCII on purpose: the script stays readable no
#  matter what LANG is set to (a non-UTF-8 locale would turn Chinese literals
#  into mojibake). Chinese messages come from the Python program instead.
# ============================================================
set -u

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd -- "$SCRIPT_DIR/.." && pwd)"
VENV_DIR="$PROJECT_ROOT/.venv"
WHEEL_DIR="$PROJECT_ROOT/wheels"
CACHE_DIR="$PROJECT_ROOT/cache"
PYTHON_PATH_FILE="$CACHE_DIR/python_path.txt"
REQUIREMENTS="$PROJECT_ROOT/requirements.txt"

# minimal supported Python, written as a version tuple for the check below
MIN_PYTHON_TUPLE="(3, 9)"
MIN_PYTHON_TEXT="3.9"

ALLOW_INSTALL=1
EXIT_OK=0
EXIT_FAIL=1

BASE_PYTHON=""      # interpreter found on this machine
PACKAGE_MANAGER=""  # apt-get / dnf / yum / apk / zypper / pacman

log()  { printf '[INFO] %s\n' "$*"; }
warn() { printf '[WARN] %s\n' "$*"; }
err()  { printf '[ERROR] %s\n' "$*" >&2; }

usage() {
    sed -n '2,30p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
}

for arg in "$@"; do
    case "$arg" in
        -h|--help) usage; exit $EXIT_OK ;;
        --no-install) ALLOW_INSTALL=0 ;;
        *) err "unknown option: $arg  (try --help)"; exit 2 ;;
    esac
done

# ------------------------------------------------------------
# helpers
# ------------------------------------------------------------

have() { command -v "$1" >/dev/null 2>&1; }

# run a command as root: directly when we already are root, else via sudo
run_privileged() {
    if [ "$(id -u)" -eq 0 ]; then
        "$@"
        return $?
    fi
    if have sudo; then
        sudo "$@"
        return $?
    fi
    return 127
}

can_install() {
    [ "$ALLOW_INSTALL" -eq 1 ] || return 1
    [ "$(id -u)" -eq 0 ] && return 0
    have sudo
}

detect_package_manager() {
    local pm
    for pm in apt-get dnf yum apk zypper pacman; do
        if have "$pm"; then
            printf '%s' "$pm"
            return 0
        fi
    done
    printf ''
}

# install_packages <package...>   (uses the global $PACKAGE_MANAGER)
install_packages() {
    [ "$#" -gt 0 ] || return 1
    [ -n "$PACKAGE_MANAGER" ] || return 127

    if ! can_install; then
        warn "no root/sudo - cannot install: $*"
        return 127
    fi

    log "installing system packages via $PACKAGE_MANAGER: $*"
    case "$PACKAGE_MANAGER" in
        apt-get)
            # a stale package list is the most common reason for a failed
            # install on a fresh container, so refresh it once beforehand
            run_privileged apt-get update -qq >/dev/null 2>&1 || true
            run_privileged apt-get install -y --no-install-recommends "$@"
            ;;
        dnf|yum) run_privileged "$PACKAGE_MANAGER" install -y "$@" ;;
        apk)     run_privileged apk add --no-cache "$@" ;;
        zypper)  run_privileged zypper --non-interactive install "$@" ;;
        pacman)  run_privileged pacman -S --noconfirm --needed "$@" ;;
        *)       return 127 ;;
    esac
}

# a usable interpreter exists, is >= MIN_PYTHON and can import venv
python_is_usable() {
    local py="$1"
    [ -n "$py" ] || return 1
    have "$py" || [ -x "$py" ] || return 1
    "$py" -c "import sys, venv; sys.exit(0 if sys.version_info >= $MIN_PYTHON_TUPLE else 1)" \
        >/dev/null 2>&1
}

# sets BASE_PYTHON (empty when nothing suitable was found)
find_python() {
    local name version candidate
    BASE_PYTHON=""
    for name in python3 python3.13 python3.12 python3.11 python3.10 python3.9 python; do
        if python_is_usable "$name"; then
            BASE_PYTHON="$(command -v "$name")"
            return 0
        fi
    done
    for version in 3.13 3.12 3.11 3.10 3.9; do
        candidate="/usr/bin/python$version"
        if [ -x "$candidate" ] && python_is_usable "$candidate"; then
            BASE_PYTHON="$candidate"
            return 0
        fi
    done
    return 1
}

python_version_of() {
    "$1" -c 'import platform; print(platform.python_version())' 2>/dev/null
}

# ------------------------------------------------------------
# 1. Python 3
# ------------------------------------------------------------
# NOTE: the ensure_* functions never print the value they produce (logs go to
# stdout as well, so capturing their output with $(...) would mix both up).
# The resolved interpreter is returned through $BASE_PYTHON instead.
ensure_python() {
    if find_python; then
        log "python found: $BASE_PYTHON ($(python_version_of "$BASE_PYTHON"))"
        return 0
    fi

    PACKAGE_MANAGER="$(detect_package_manager)"
    if [ -z "$PACKAGE_MANAGER" ]; then
        err "no python3 >= $MIN_PYTHON_TEXT found and no known package manager available."
        err "install Python $MIN_PYTHON_TEXT or newer by hand, then run this script again."
        return 1
    fi

    warn "python3 >= $MIN_PYTHON_TEXT not found (or the venv module is missing) - installing it"
    # venv / pip usually live in separate packages, so all of them are requested
    case "$PACKAGE_MANAGER" in
        apt-get) install_packages python3 python3-venv python3-pip ;;
        apk)     install_packages python3 py3-pip py3-virtualenv ;;
        pacman)  install_packages python python-pip ;;
        *)       install_packages python3 python3-pip ;;
    esac

    if find_python; then
        log "python installed: $BASE_PYTHON ($(python_version_of "$BASE_PYTHON"))"
        return 0
    fi

    err "python3 is still unavailable. Run this by hand and retry:"
    case "$PACKAGE_MANAGER" in
        apt-get) err "    sudo apt-get install -y python3 python3-venv python3-pip" ;;
        dnf|yum) err "    sudo $PACKAGE_MANAGER install -y python3 python3-pip" ;;
        apk)     err "    sudo apk add python3 py3-pip py3-virtualenv" ;;
        zypper)  err "    sudo zypper install python3 python3-pip python3-virtualenv" ;;
        pacman)  err "    sudo pacman -S python python-pip" ;;
        *)       err "    install python$MIN_PYTHON_TEXT or newer with your package manager" ;;
    esac
    return 1
}

# ------------------------------------------------------------
# 2. virtual environment
# ------------------------------------------------------------
venv_python_is_usable() {
    [ -x "$VENV_DIR/bin/python" ] || return 1
    "$VENV_DIR/bin/python" -c 'import sys' >/dev/null 2>&1
}

# a venv without pip is useless for us - Debian/Ubuntu ship ensurepip in a
# separate package (python3-venv), so "python3 -m venv" silently produces a
# pip-less environment there
venv_pip_works() {
    venv_python_is_usable || return 1
    "$VENV_DIR/bin/python" -m pip --version >/dev/null 2>&1
}

create_venv_once() {
    rm -rf "$VENV_DIR"
    if "$BASE_PYTHON" -m venv "$VENV_DIR" >/dev/null 2>&1 && venv_pip_works; then
        return 0
    fi
    # retry without the pip bootstrap, then add pip with ensurepip
    rm -rf "$VENV_DIR"
    "$BASE_PYTHON" -m venv --without-pip "$VENV_DIR" >/dev/null 2>&1 || return 1
    venv_python_is_usable || return 1
    "$VENV_DIR/bin/python" -m ensurepip --upgrade >/dev/null 2>&1 || true
    venv_pip_works
}

ensure_venv() {
    if venv_pip_works; then
        log "virtual environment reused: $VENV_DIR"
        return 0
    fi
    if [ -e "$VENV_DIR" ]; then
        # a .venv copied from another machine keeps absolute paths to the
        # original interpreter in pyvenv.cfg and bin/, and one created without
        # ensurepip has no pip at all - both are rebuilt here
        warn "existing .venv is unusable here (copied from another machine, or without pip) - rebuilding"
        rm -rf "$VENV_DIR"
    fi

    log "creating virtual environment: $VENV_DIR"
    if create_venv_once; then
        log "virtual environment created"
        return 0
    fi

    # most common reason: ensurepip is missing (python3-venv not installed)
    warn "the virtual environment would have no pip - installing the venv/pip packages"
    if [ -z "$PACKAGE_MANAGER" ]; then
        PACKAGE_MANAGER="$(detect_package_manager)"
    fi
    case "$PACKAGE_MANAGER" in
        apt-get) install_packages python3-venv python3-pip >/dev/null 2>&1 || true ;;
        apk)     install_packages py3-pip py3-virtualenv >/dev/null 2>&1 || true ;;
        pacman)  install_packages python-pip >/dev/null 2>&1 || true ;;
        zypper)  install_packages python3-pip python3-virtualenv >/dev/null 2>&1 || true ;;
        dnf|yum) install_packages python3-pip >/dev/null 2>&1 || true ;;
        '')      : ;;
        *)       install_packages python3-pip >/dev/null 2>&1 || true ;;
    esac

    if create_venv_once; then
        log "virtual environment created (pip came from the distribution packages)"
        return 0
    fi

    err "could not create a virtual environment with pip at $VENV_DIR"
    err "run this by hand and try again:"
    err "    Debian/Ubuntu : sudo apt-get install -y python3-venv python3-pip"
    err "    RHEL/CentOS   : sudo dnf install -y python3-pip"
    err "    Alpine        : sudo apk add py3-pip py3-virtualenv"
    return 1
}

# ------------------------------------------------------------
# 3. python dependencies
# ------------------------------------------------------------
ensure_dependencies() {
    local py="$VENV_DIR/bin/python"
    if "$py" -c 'import openpyxl, requests' >/dev/null 2>&1; then
        log "dependencies already installed (openpyxl, requests)"
        return 0
    fi
    if ! "$py" -m pip --version >/dev/null 2>&1; then
        err "this virtual environment has no pip - re-run scripts/setup_env.sh"
        return 1
    fi

    if [ -d "$WHEEL_DIR" ] && ls "$WHEEL_DIR"/*.whl >/dev/null 2>&1; then
        log "installing dependencies from local wheels (offline mode)"
        if "$py" -m pip install --no-index --find-links "$WHEEL_DIR" -r "$REQUIREMENTS" \
            >/dev/null 2>&1; then
            log "dependencies installed from wheels/"
            return 0
        fi
        warn "installing from wheels/ failed (python version or platform mismatch?) - using PyPI"
    fi

    log "installing dependencies from PyPI (needs network access)"
    "$py" -m pip install --upgrade pip >/dev/null 2>&1 || true
    if ! "$py" -m pip install -r "$REQUIREMENTS"; then
        err "could not install the dependencies. Check the network / proxy, or place"
        err "matching .whl files into $WHEEL_DIR and run this script again."
        return 1
    fi
    log "dependencies installed"
    return 0
}

# ------------------------------------------------------------
# 4. dig
# ------------------------------------------------------------
ensure_dig() {
    if have dig; then
        log "dig found: $(command -v dig)"
        return 0
    fi
    if [ -x "$PROJECT_ROOT/tools/bind/dig" ]; then
        log "dig found (portable): tools/bind/dig"
        return 0
    fi

    if [ -z "$PACKAGE_MANAGER" ]; then
        PACKAGE_MANAGER="$(detect_package_manager)"
    fi
    if [ -z "$PACKAGE_MANAGER" ]; then
        warn "dig is missing and no known package manager is available"
        return 1
    fi

    warn "dig is missing - installing it"
    case "$PACKAGE_MANAGER" in
        # Debian 13 / Ubuntu 24.04+ moved dig from 'dnsutils' to 'bind9-dnsutils'
        apt-get) install_packages dnsutils >/dev/null 2>&1 \
                     || install_packages bind9-dnsutils >/dev/null 2>&1 || true ;;
        apk)     install_packages bind-tools ;;
        pacman)  install_packages bind ;;
        # zypper and dnf/yum both call it bind-utils
        *)       install_packages bind-utils ;;
    esac

    if have dig; then
        log "dig installed: $(command -v dig)"
        return 0
    fi
    err "dig is still unavailable. Install it by hand, e.g.:"
    err "    Debian/Ubuntu : sudo apt-get install -y dnsutils   (or bind9-dnsutils)"
    err "    RHEL/CentOS   : sudo dnf install -y bind-utils"
    err "    Alpine        : sudo apk add bind-tools"
    err "    or copy a working dig into $PROJECT_ROOT/tools/bind/"
    return 1
}

# ------------------------------------------------------------
# 5. ICMP permission check (only a hint - the program still runs)
# ------------------------------------------------------------
can_send_icmp() {
    local ping_bin
    ping_bin="$(command -v ping 2>/dev/null || true)"
    [ -n "$ping_bin" ] || return 1
    [ "$(id -u)" -eq 0 ] && return 0
    if have getcap && getcap "$ping_bin" 2>/dev/null | grep -q cap_net_raw; then
        return 0
    fi
    if [ -u "$ping_bin" ]; then          # setuid bit
        return 0
    fi
    local range lo hi gid
    range="$(cat /proc/sys/net/ipv4/ping_group_range 2>/dev/null || true)"
    if [ -n "$range" ]; then
        set -- $range
        lo="${1:-0}"
        hi="${2:-0}"
        gid="$(id -g 2>/dev/null || echo 0)"
        if [ "$lo" -le "$gid" ] 2>/dev/null && [ "$gid" -le "$hi" ] 2>/dev/null; then
            return 0
        fi
    fi
    return 1
}

warn_about_icmp() {
    if can_send_icmp; then
        log "ICMP check: the current user can send pings"
        return 0
    fi
    warn "this user may not be allowed to send ICMP packets; the latency and loss"
    warn "columns would stay empty. One-time fix (pick either line):"
    warn "    sudo setcap cap_net_raw+ep \"\$(command -v ping)\""
    warn "    sudo sysctl -w net.ipv4.ping_group_range='0 2147483647'"
    return 1
}

# the same check under its reporting name used by main()
report_icmp() { warn_about_icmp; }

# ------------------------------------------------------------
# 6. remember the interpreter for the launcher
# ------------------------------------------------------------
write_cache() {
    local py="$1" host
    host="$(uname -n 2>/dev/null || printf '')"
    mkdir -p "$CACHE_DIR" || return 1
    {
        printf 'machine=%s\n' "$host"
        printf 'python=%s\n' "$py"
    } > "$PYTHON_PATH_FILE" || return 1
    log "interpreter recorded in cache/python_path.txt"
    log "    machine=$host"
    log "    python=$py"
}

# ------------------------------------------------------------
# main
# ------------------------------------------------------------
main() {
    log "project : $PROJECT_ROOT"
    if [ "$ALLOW_INSTALL" -eq 0 ]; then
        log "mode    : diagnose only (--no-install)"
    fi

    ensure_python || return $EXIT_FAIL
    ensure_venv || return $EXIT_FAIL
    ensure_dependencies || return $EXIT_FAIL

    local status=$EXIT_OK
    ensure_dig || status=$EXIT_FAIL
    report_icmp || true
    write_cache "$VENV_DIR/bin/python" || true

    if [ "$status" -eq "$EXIT_OK" ]; then
        log "environment ready. Start the test with:  ./run_dns_test.sh"
    else
        err "environment is incomplete - fix the errors above and run this script again."
    fi
    return "$status"
}

main
exit $?
