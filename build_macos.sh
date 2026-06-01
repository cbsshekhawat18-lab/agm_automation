#!/bin/bash
# Local macOS build.  Produces dist/AGM_Final_Set (~14MB) + a release .zip.
#
# Prereqs: Python 3.13 with python-docx, openpyxl, pyinstaller installed.
#   pip install -r requirements.txt pyinstaller
#
# Usage:
#   ./build_macos.sh           # build + zip
#   ./build_macos.sh --clean   # also wipe build/ + dist/ first
set -euo pipefail
cd "$(dirname "$0")"

if [ "${1:-}" = "--clean" ]; then
    echo ">> Cleaning build/ and dist/"
    rm -rf build dist
fi

echo ">> Building demo Master_Input"
python3 sample/build_demo.py

echo ">> Running PyInstaller"
python3 -m PyInstaller agm_final_set.spec --noconfirm

VERSION="${VERSION:-dev}"
ZIPNAME="AGM_Final_Set-macos-${VERSION}.zip"
STAGE="dist/AGM_Final_Set-${VERSION}-macos"

echo ">> Staging release in ${STAGE}"
rm -rf "$STAGE" "dist/${ZIPNAME}"
mkdir -p "$STAGE"
cp dist/AGM_Final_Set "$STAGE/"
cp README.md "$STAGE/"
cp LICENSE "$STAGE/"
[ -f USAGE.md ]   && cp USAGE.md   "$STAGE/"
[ -f PRIVACY.md ] && cp PRIVACY.md "$STAGE/"
# Both templates ship side-by-side — user renames whichever they want to use.
cp sample/Master_Input_DEMO.xlsx "$STAGE/"
cp sample/Master_Input_EMPTY.xlsx "$STAGE/"

echo ">> Zipping"
cd dist && zip -r "$ZIPNAME" "AGM_Final_Set-${VERSION}-macos" >/dev/null
echo ">> Done: dist/${ZIPNAME}"
