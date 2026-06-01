#!/bin/bash
# Double-click this file (Mac) to generate the AGM Final Set.
# Requires Python 3.10+ with python-docx and openpyxl installed.

cd "$(dirname "$0")"

# Pick the first Python that (a) exists, (b) loads stdlib cleanly, and
# (c) ideally already has python-docx + openpyxl. Tested in this order so
# we prefer a known-stable LTS over a possibly-broken newest install.
PY=""
for candidate in python3.13 python3.12 python3.11 python3.10 python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
        # Sanity-check the interpreter can import its own stdlib.
        if "$candidate" -c "import sys, sysconfig" >/dev/null 2>&1; then
            PY="$candidate"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo
    echo "ERROR: Python is not installed (or the installed Python is broken)."
    echo "Install Python 3 from https://www.python.org/downloads/ then re-run this file."
    echo
    read -n 1 -s -r -p "Press any key to close..."
    exit 1
fi

echo "Using interpreter: $($PY --version 2>&1) at $(command -v $PY)"
echo

if ! "$PY" -c "import docx, openpyxl" >/dev/null 2>&1; then
    echo "Installing required libraries (python-docx, openpyxl)..."
    "$PY" -m pip install --quiet python-docx openpyxl || {
        echo
        echo "ERROR: Failed to install python-docx and openpyxl."
        echo "Run this manually:  $PY -m pip install python-docx openpyxl"
        echo
        read -n 1 -s -r -p "Press any key to close..."
        exit 1
    }
fi

"$PY" generate_final_set.py
echo
read -n 1 -s -r -p "Press any key to close..."
