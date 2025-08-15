# Graph Generator — Multi‑Cono (Codespaces Only)

This package rebuilds the final graph workflow we agreed on:

- Auto‑discovers the **latest** consolidated workbook per `./Cono*/Consolidated_reports`.
- Cleans **Section 11** and writes to `Modified_Section11`.
- Builds charts from `config_grph.json` using **dynamic multi‑series** by `y_suffix` where headers follow `MMDDYYYY_<suffix>`.
- Writes a **new** output workbook with sheets ordered: **Graph_*** → **Modified_Section11** → **Original_*** → **Run_Log**.

> No `--config` needed. The script reads `./config_grph.json` automatically.

## Quick Start (Codespaces)

```bash
pip install -r requirements.txt  # if needed
python inspect_latest_wb.py      # shows which file will be picked per Cono
python gen_grph.py               # full run
python gen_grph.py --dryrun      # discovery only
```

### Expected Input Layout

```
./Cono1/Consolidated_reports/*.xlsx
./Cono2/Consolidated_reports/*.xlsx
./Cono3/Consolidated_reports/*.xlsx
```

Files with `_graph`, `_graph_`, or `_Grph` in the name are ignored. 
Filenames like `Consolidate_report_08072025_16_08.xlsx` are preferred (the timestamp in the name is used to pick latest). If that pattern is missing, file **modification time** is used.

### Chart Rules

- Each chart in `config_grph.json` specifies:
  - `sheet`: source worksheet name
  - `x_col`: category column (e.g., `TakenBy`, `Oper`)
  - `y_suffix`: metric suffix to gather (e.g., `OE_Count`)
  - `chart_type`: one of `line`, `column`, `bar`, `scatter`, `stacked_column`
  - `title`: Excel chart title
- The script finds all columns matching `^\d{8}_<y_suffix>$`, **sorts by date** ascending, coerces to numeric, and drops rows where all Y values are empty.
- Each chart is written to its own `Graph_<...>` sheet, with human‑friendly series labels (`MM/DD/YYYY`).

### Section 11 Cleanup

- If a sheet named any of `Section11`, `Section 11`, `Sec11`, `Section_11` exists:
  - Trims whitespace
  - Drops fully empty rows/cols
  - Removes repeated header rows that appear mid‑sheet
  - Writes the result to `Modified_Section11`

### Output File Name

Pattern is controlled by `output.filename_pattern`:
```
{cono}_Src_{src_ts}__Grph_{run_ts_EST}.xlsx
```
- `{src_ts}` is taken from the source file name if present, otherwise from file modified time.
- `{run_ts_EST}` uses America/New_York.

### Requirements

- Python 3.12
- pandas
- openpyxl

See `requirements.txt`.

## Notes / Limitations

- Axis tick label rotation is limited in `openpyxl`; defaults are used.
- Original data sheets are copied as values only to `Original_*` tabs.
- If a chart's `sheet` or `x_col` is missing, it is **skipped** and a note is written to `Run_Log`.
