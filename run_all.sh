#!/usr/bin/env bash
# One-command run: (re)generate raw data, clean it, rebuild all reports.
set -e
python src/generate_data.py   # skip this line if you use your own raw CSV
python src/pipeline.py
