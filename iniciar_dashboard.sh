#!/usr/bin/env bash
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet
python3 dashboard.py
