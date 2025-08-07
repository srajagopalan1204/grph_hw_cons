import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
import openpyxl
import json
import os
import glob
from datetime import datetime
import sys

# Handle --dryrun argument
dryrun = "--dryrun" in sys.argv

# Setup logging directory and file
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_path = os.path.join("logs", f"log_{timestamp}.txt")

def log(message):
    print(message)
    with open(log_path, "a") as log_file:
        log_file.write(message + "\n")

# Clear previous log content
open(log_path, "w").close()

# Load configuration
with open("config_grph.json", "r") as f:
    config = json.load(f)

# Known Conos
known_conos = ["Cono1", "Cono2", "Cono3"]

for cono in known_conos:
    folder = f"./{cono}/Consolidated_reports/"
    os.makedirs(folder, exist_ok=True)

    # Find latest file excluding '_grph'
    files = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if "_graph" not in os.path.basename(f)]
    if not files:
        log(f"[WARNING] No valid Excel reports found in {folder}")
        continue
    file_path = max(files, key=os.path.getmtime)
    log(f"[INFO] Processing workbook for {cono}: {file_path}")

    try:
        wb = load_workbook(file_path)
    except Exception as e:
        log(f"[ERROR] Could not load workbook: {e}")
        continue

    base_file_name = os.path.splitext(os.path.basename(file_path))[0]
    time_label = datetime.now().strftime("%Y%m%d_%H%M")
    output_file_name = f"{cono}_{base_file_name}_graph_{time_label}.xlsx"
    output_file_path = os.path.join(folder, output_file_name)

    # Get chart config
    wb_def = next((w for w in config.get("workbooks", []) if w.get("cono") == cono), None)
    if not wb_def:
        log(f"[WARNING] No chart configuration found for {cono}. Skipping...")
        continue

    charts = wb_def.get("charts", [])
    if not isinstance(charts, list) or not charts:
        log(f"[WARNING] No chart entries found for {cono}. Skipping...")
        continue

    for chart_def in charts:
        sheet_name = chart_def.get("sheet")
        output_sheet = chart_def.get("output_sheet")
        x_col = chart_def.get("x_col")
        title = chart_def.get("title", f"{cono}_{output_sheet}")

        # Load sheet into DataFrame
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            log(f"[ERROR] Could not read sheet {sheet_name}: {e}")
            continue

        if len(df) < 2:
            log(f"[INFO] Sheet '{sheet_name}' has fewer than 2 rows. Skipping chart.")
            continue

        if x_col not in df.columns:
            log(f"[WARNING] x_col '{x_col}' not found in {sheet_name}. Skipping...")
            continue

        # Dynamically find all Y columns (excluding x_col)
        y_cols = [col for col in df.columns if col != x_col and isinstance(col, str) and "_" in col]

        if not y_cols:
            log(f"[INFO] No valid Y columns found in {sheet_name}. Skipping chart.")
            continue

        # Remove output sheet if exists
        if output_sheet in wb.sheetnames:
            del wb[output_sheet]
        ws = wb.create_sheet(output_sheet, 0)

        headers = [x_col] + y_cols
        for c_idx, col_name in enumerate(headers, start=1):
            ws.cell(row=1, column=c_idx, value=col_name)

        for r_idx, row in enumerate(df[headers].values.tolist(), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Create line chart
        chart = LineChart()
        chart.title = title
        chart.y_axis.title = "Count"
        chart.x_axis.title = x_col
        chart.x_axis.majorTickMark = "in"
        chart.y_axis.majorTickMark = "in"
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.label_rotation = -45

        data = Reference(ws, min_col=2, min_row=1, max_col=1 + len(y_cols), max_row=len(df) + 1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=len(df) + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        ws.add_chart(chart, "E2")
        log(f"[SUCCESS] Chart created: {output_sheet} in {cono}")

    if dryrun:
        log(f"[DRYRUN] Skipped saving {output_file_path}")
    else:
        try:
            wb.save(output_file_path)
            log(f"[SAVED] Graphs saved to {output_file_path}")
        except Exception as e:
            log(f"[ERROR] Could not save output: {e}")
