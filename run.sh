#!/bin/bash
set -e
cd ~/german-poetry-v5

echo "=== STEP 1: Extract poems from TEI XML ==="
python3 scripts/01_extract.py

echo ""
echo "=== STEP 2: Compute token-level surprisal/entropy ==="
python3 scripts/02_compute.py

echo ""
echo "=== STEP 3: Build Excel workbook ==="
python3 scripts/03_make_workbook.py

echo ""
echo "=== DONE ==="
echo "Output files:"
ls -lh ~/german-poetry-v5/output/

echo ""
echo "Model used: dbmdz/german-gpt2"
