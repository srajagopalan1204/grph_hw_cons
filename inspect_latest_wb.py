#!/usr/bin/env python3
"""
inspect_latest_wb.py
Quickly shows which latest .xlsx would be chosen per ./Cono*/Consolidated_reports
Excludes files containing: _graph, _graph_, _Grph
"""
import os, glob, re, datetime as dt

IGNORES = ["_graph", "_graph_", "_Grph"]

def _is_ignored(filename: str) -> bool:
    if not filename.lower().endswith(".xlsx"):
        return True
    return any(s.lower() in filename.lower() for s in IGNORES)

def _parse_timestamp_from_name(fname: str):
    m = re.search(r'(\d{8})_(\d{2})_(\d{2})', fname)
    if not m:
        return None
    try:
        return dt.datetime.strptime(m.group(1)+m.group(2)+m.group(3), "%m%d%Y%H%M")
    except Exception:
        return None

def pick_latest_xlsx(path: str):
    files = [p for p in glob.glob(os.path.join(path, "*.xlsx")) if not _is_ignored(os.path.basename(p))]
    if not files:
        return None
    scored = []
    for p in files:
        t = _parse_timestamp_from_name(os.path.basename(p))
        if t:
            scored.append((t.timestamp(), p))
        else:
            scored.append((os.path.getmtime(p), p))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]

def main():
    cono_paths = sorted(glob.glob("./Cono*/Consolidated_reports"))
    if not cono_paths:
        print("No Cono paths found.")
        return
    for path in cono_paths:
        latest = pick_latest_xlsx(path)
        print(f"{path} -> {latest if latest else 'No eligible .xlsx'}")

if __name__ == "__main__":
    main()
