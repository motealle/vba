#!/usr/bin/env python3
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
errors = []
public_names = {}

for path in sorted(ROOT.rglob("*.bas")):
    data = path.read_bytes()
    try:
        text = data.decode("ascii")
    except UnicodeDecodeError as exc:
        errors.append(f"{path.relative_to(ROOT)}: non-ASCII byte at offset {exc.start}")
        continue

    if not re.search(r"(?im)^\s*Option\s+Explicit\s*$", text):
        errors.append(f"{path.relative_to(ROOT)}: missing Option Explicit")

    for m in re.finditer(r"(?im)^\s*Public\s+(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)", text):
        name = m.group(1).lower()
        public_names.setdefault(name, []).append(str(path.relative_to(ROOT)))

    for line_no, line in enumerate(text.splitlines(), 1):
        if "Tested on" in line and not re.search(r"Tested on \d{4}-\d{2}-\d{2}", line):
            errors.append(f"{path.relative_to(ROOT)}:{line_no}: Tested on comment needs YYYY-MM-DD")

for name, paths in public_names.items():
    if len(paths) > 1:
        errors.append(f"duplicate public procedure {name}: {', '.join(paths)}")

if errors:
    print("VBA static checks failed:")
    for error in errors:
        print("-", error)
    sys.exit(1)

print("VBA static checks passed.")
print(f"Checked {len(list(ROOT.rglob('*.bas')))} BAS modules.")
