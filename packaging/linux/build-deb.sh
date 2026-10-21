#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/../.." && pwd)
DIST_DIR=${CHROMEAUTORUN_DIST_DIR:-$ROOT/dist/linux}
VERSION=${CHROMEAUTORUN_VERSION:-1.0.2}
PYTHON_VERSION=${CHROMEAUTORUN_PYTHON_VERSION:-3.12.14}
STANDALONE_TAG=${CHROMEAUTORUN_STANDALONE_TAG:-20260924}
CHROMEDRIVER_VERSION=${CHROMEAUTORUN_CHROMEDRIVER_VERSION:-151.0.7922.108}
MIRROR=${CHROMEAUTORUN_STANDALONE_MIRROR:-https://mirror.nju.edu.cn/github-release/astral-sh/python-build-standalone}
FALLBACK=${CHROMEAUTORUN_STANDALONE_FALLBACK:-https://github.com/astral-sh/python-build-standalone/releases/download}

for tool in dpkg dpkg-deb tar sha256sum sed find chmod; do
    command -v "$tool" >/dev/null 2>&1 || {
        echo "[error] Required build tool not found: $tool" >&2
        exit 1
    }
done

ARCH=$(dpkg --print-architecture)
case "$ARCH" in
    amd64) STANDALONE_ARCH=x86_64; CHROMEDRIVER_PLATFORM=linux64 ;;
    arm64) STANDALONE_ARCH=aarch64; CHROMEDRIVER_PLATFORM=linux-arm64 ;;
    *)
        echo "[error] Unsupported package architecture: $ARCH (expected amd64 or arm64)" >&2
        exit 1
        ;;
esac

ASSET="cpython-${PYTHON_VERSION}+${STANDALONE_TAG}-${STANDALONE_ARCH}-unknown-linux-gnu-install_only.tar.gz"
CACHE_BASE=${XDG_CACHE_HOME:-$HOME/.cache}
CACHE_DIR=$CACHE_BASE/chrome-autorun/python-runtime
ARCHIVE=${CHROMEAUTORUN_PYTHON_ARCHIVE:-$CACHE_DIR/$ASSET}
mkdir -p "$CACHE_DIR" "$DIST_DIR"

if [ ! -f "$ARCHIVE" ]; then
    FETCH=
    for candidate in curl wget; do
        if command -v "$candidate" >/dev/null 2>&1; then
            FETCH=$candidate
            break
        fi
    done
    [ -n "$FETCH" ] || { echo '[error] Install curl or wget to download bundled Python.' >&2; exit 1; }

    ASSET_ENCODED=$(printf '%s' "$ASSET" | sed 's/+/%2B/g')
    DOWNLOADED=0
    for url in "$MIRROR/$STANDALONE_TAG/$ASSET_ENCODED" "$FALLBACK/$STANDALONE_TAG/$ASSET"; do
        echo "[python] Downloading $url"
        if [ "$FETCH" = curl ]; then
            if curl -fL --retry 2 --connect-timeout 30 -o "$ARCHIVE.part" "$url"; then DOWNLOADED=1; break; fi
        else
            if wget --tries=2 --timeout=60 -O "$ARCHIVE.part" "$url"; then DOWNLOADED=1; break; fi
        fi
    done
    [ "$DOWNLOADED" -eq 1 ] || { rm -f "$ARCHIVE.part"; echo '[error] Python runtime download failed.' >&2; exit 1; }
    mv "$ARCHIVE.part" "$ARCHIVE"
fi

WORK_DIR=$(mktemp -d "${TMPDIR:-/tmp}/chrome-autorun-deb.XXXXXXXX")
trap 'rm -rf -- "$WORK_DIR"' EXIT HUP INT TERM
RUNTIME_EXTRACT=$WORK_DIR/runtime-extract
PACKAGE_ROOT=$WORK_DIR/package
APP_ROOT=$PACKAGE_ROOT/opt/chrome-autorun
mkdir -p "$RUNTIME_EXTRACT" "$APP_ROOT/defaults" "$APP_ROOT/vendor" \
    "$APP_ROOT/runtime" "$APP_ROOT/drivers" "$PACKAGE_ROOT/usr/bin" "$PACKAGE_ROOT/usr/share/applications" \
    "$PACKAGE_ROOT/usr/share/doc/chrome-autorun" "$PACKAGE_ROOT/DEBIAN"

echo '[python] Extracting the portable CPython runtime.'
tar -xzf "$ARCHIVE" -C "$RUNTIME_EXTRACT"
BUILD_PY=$RUNTIME_EXTRACT/python/bin/python3
[ -x "$BUILD_PY" ] || { echo '[error] Bundled Python archive has an unexpected layout.' >&2; exit 1; }

CHROMEDRIVER_URL="https://storage.googleapis.com/chrome-for-testing-public/$CHROMEDRIVER_VERSION/$CHROMEDRIVER_PLATFORM/chromedriver-$CHROMEDRIVER_PLATFORM.zip"

DRIVER_FETCH=
for candidate in curl wget; do
    if command -v "$candidate" >/dev/null 2>&1; then
        DRIVER_FETCH=$candidate
        break
    fi
done
[ -n "$DRIVER_FETCH" ] || { echo '[error] Install curl or wget to download bundled ChromeDriver.' >&2; exit 1; }
DRIVER_ARCHIVE=$WORK_DIR/chromedriver.zip
echo "[driver] Downloading ChromeDriver $CHROMEDRIVER_VERSION ($CHROMEDRIVER_PLATFORM)."
if [ "$DRIVER_FETCH" = curl ]; then
    curl -fL --retry 2 --connect-timeout 30 -o "$DRIVER_ARCHIVE" "$CHROMEDRIVER_URL"
else
    wget --tries=2 --timeout=60 -O "$DRIVER_ARCHIVE" "$CHROMEDRIVER_URL"
fi
"$BUILD_PY" - "$DRIVER_ARCHIVE" "$APP_ROOT/drivers/chromedriver" <<'PY'
import sys
from pathlib import Path
from zipfile import ZipFile

archive, destination = map(Path, sys.argv[1:])
with ZipFile(archive) as bundle:
    candidates = [name for name in bundle.namelist() if Path(name).name == 'chromedriver' and not name.endswith('/')]
    if len(candidates) != 1:
        raise SystemExit(f'Expected one chromedriver binary in archive, found {len(candidates)}')
    destination.write_bytes(bundle.read(candidates[0]))
destination.chmod(0o755)
PY
CHROMEDRIVER_BUILD_INFO=$("$APP_ROOT/drivers/chromedriver" --version)
echo "[driver] $CHROMEDRIVER_BUILD_INFO"

echo '[deps] Installing Selenium and openpyxl into the package.'
"$BUILD_PY" -m pip install --disable-pip-version-check --no-compile \
    --target "$APP_ROOT/vendor" -r "$ROOT/requirements-linux.txt"
PYTHONPATH="$APP_ROOT/vendor" "$BUILD_PY" -c \
    'import selenium, openpyxl; print("Selenium", selenium.__version__, "openpyxl", openpyxl.__version__)'
"$BUILD_PY" -m py_compile "$ROOT/main.py"

echo '[package] Assembling DEB filesystem.'
cp -R "$RUNTIME_EXTRACT/python" "$APP_ROOT/runtime/"
cp "$ROOT/main.py" "$APP_ROOT/main.py"
cp "$ROOT/config.json" "$APP_ROOT/defaults/config.json"
cp "$ROOT/config.xlsx" "$APP_ROOT/defaults/config.xlsx"
cp "$SCRIPT_DIR/chrome-autorun" "$PACKAGE_ROOT/usr/bin/chrome-autorun"
cp "$SCRIPT_DIR/chrome-autorun.desktop" "$PACKAGE_ROOT/usr/share/applications/chrome-autorun.desktop"
chmod 755 "$PACKAGE_ROOT/usr/bin/chrome-autorun"

cat > "$PACKAGE_ROOT/usr/share/doc/chrome-autorun/build-info.txt" <<INFO
Package: chrome-autorun
Version: $VERSION
Architecture: $ARCH
Bundled Python: $PYTHON_VERSION (python-build-standalone)
Python runtime archive: $ASSET
Application: ChromeAutoRun
Chrome browser: supplied by the target system
Bundled ChromeDriver: $CHROMEDRIVER_BUILD_INFO
ChromeDriver source: Chrome for Testing
INFO

cat > "$PACKAGE_ROOT/DEBIAN/control" <<CONTROL
Package: chrome-autorun
Version: $VERSION
Section: net
Priority: optional
Architecture: $ARCH
Depends: libc6 (>= 2.17), libstdc++6, xdg-utils
Maintainer: ChromeAutoRun
Description: Chinese-language Selenium Chrome report automation tool
 ChromeAutoRun automates FineReport report queries and exports.
 The Chrome or Chromium browser must be installed separately.
 ChromeDriver is bundled; use a build matching the target browser version for offline operation.
CONTROL

find "$PACKAGE_ROOT" -type d -exec chmod 755 {} +
find "$PACKAGE_ROOT" -type f ! -path '*/usr/bin/chrome-autorun' -exec chmod 644 {} +
chmod 755 "$PACKAGE_ROOT/usr/bin/chrome-autorun"
chmod 755 "$APP_ROOT/drivers/chromedriver"
find "$APP_ROOT/runtime/python/bin" -type f -exec chmod 755 {} +
find "$APP_ROOT/vendor" -type f -name 'selenium-manager' -exec chmod 755 {} +

DEB=$DIST_DIR/chrome-autorun_${VERSION}_${ARCH}.deb
dpkg-deb --build --root-owner-group "$PACKAGE_ROOT" "$DEB"
dpkg-deb --info "$DEB" >/dev/null
(cd "$DIST_DIR" && sha256sum "${DEB##*/}" > "${DEB##*/}.sha256")
echo "[done] $DEB"
