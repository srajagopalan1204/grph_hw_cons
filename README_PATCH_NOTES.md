# Graph Generator — Patched Notes

**Patches applied (2025-08-10):**
- Fixed `TypeError` by setting `chart.legend = Legend()` (openpyxl requires a `Legend` object, not `True`).
- Auto-truncate/uniquify sheet names to avoid the 31-character Excel limit.

Usage remains the same:
```bash
python inspect_latest_wb.py
python gen_grph.py
```
