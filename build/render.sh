#!/bin/bash
# usage: render.sh file.pptx pages(1,2,5)
set -e
f="$1"; pages="$2"
python3 - "$f" "$pages" <<'PY'
import sys, subprocess, os
PY
