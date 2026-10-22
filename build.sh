#!/usr/bin/env bash
# Prepare the static site: copy the converter and download the pure-Python wheels it needs.
# (Pyodide itself provides cryptography, Pillow and charset-normalizer.)
# Usage: bash build.sh      then publish this folder (everything except tests/) to GitHub Pages.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
PIP="${PIP:-$HERE/../coa_webapp/.venv/bin/pip}"
cd "$HERE"

cp ../coa_to_excel/coa_to_excel.py py/coa_to_excel.py
rm -rf wheels && mkdir wheels
"$PIP" download --quiet --no-deps --only-binary=:all: --dest wheels \
    pdfplumber==0.11.10 pdfminer.six==20260107 openpyxl==3.1.5 et_xmlfile==2.0.0
python3 -c "import json,os; print(json.dumps(sorted(f for f in os.listdir('wheels') if f.endswith('.whl'))))" > wheels/manifest.json
touch .nojekyll                      # serve files as-is on GitHub Pages
ls -l wheels; du -sh "$HERE" --exclude=tests
