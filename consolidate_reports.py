
"""
Consolidate multi-Cono weekly ERP training reports (Pivot by Date_of_rep, Multi-Tab).
- Uses config.json5 to define key columns and comp columns.
- Recursively scans date folders, reads each section file.
- Extracts 'Date_of_rep' from file for column headers (<MMDDYYYY>_<CompCol>).
- Produces one workbook per Cono with one sheet per Section file.
"""

import os
import glob
import pandas as pd
from datetime import datetime
import json5

def load_config(config_path="config.json5"):
    with open(config_path, "r") as f:
        return json5.load(f)

def normalize_date(date_value):
    """Convert Date_of_rep value (MM/DD/YY) to MMDDYYYY string."""
    try:
        # Convert using pandas to_datetime
        dt = pd.to_datetime(str(date_value), errors='coerce')
        if pd.isnull(dt):
            return None
        return dt.strftime("%m%d%Y")
    except Exception:
        return None

def process_section(section_name, key_cols, comp_cols, files_by_date):
    """Merge data horizontally by key_cols, pivoting comp_cols with date prefix."""
    merged_df = pd.DataFrame()

    for date_key, file_path in files_by_date.items():
        try:
            df = pd.read_excel(file_path, dtype=str)
            # Ensure key columns exist
            if not all(col in df.columns for col in key_cols):
                continue
            # Extract Date_of_rep value
            if "Date_of_rep" in df.columns:
                rep_date = normalize_date(df["Date_of_rep"].iloc[0])
            else:
                rep_date = date_key  # fallback to folder date if missing
            if rep_date is None:
                continue
            # Select required columns
            subset = df[key_cols + comp_cols].copy()
            # Rename comp cols with date prefix
            rename_map = {col: f"{rep_date}_{col}" for col in comp_cols}
            subset = subset.rename(columns=rename_map)

            # Merge into master
            if merged_df.empty:
                merged_df = subset
            else:
                merged_df = pd.merge(merged_df, subset, on=key_cols, how="outer")
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    return merged_df

def consolidate_reports():
    config = load_config()

    for cono, details in config.items():
        source_path = details["source_path"]
        destination_path = details["destination_path"]
        files = details["files"]

        os.makedirs(destination_path, exist_ok=True)

        timestamp = datetime.now().strftime("%m%d%Y_%H_%M")
        output_filename = f"Consolidate_report_{timestamp}.xlsx"
        output_path = os.path.join(destination_path, output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            any_data = False

            # For each section defined in config
            for section_file, mapping in files.items():
                key_cols = mapping.get("key_cols", [])
                comp_cols = mapping.get("Comp_cols", [])

                # Find files across all subfolders
                pattern = os.path.join(source_path, "*", section_file)
                matches = [p for p in glob.glob(pattern) if "_del" not in p.lower()]

                # Map by folder name (date folder)
                files_by_date = {}
                for match in matches:
                    # folder name as date_key (fallback)
                    date_key = os.path.basename(os.path.dirname(match))
                    files_by_date[date_key] = match

                if files_by_date:
                    merged_df = process_section(section_file, key_cols, comp_cols, files_by_date)
                    if not merged_df.empty:
                        sheet_name = os.path.splitext(section_file)[0][:31]
                        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        any_data = True

            if any_data:
                print(f"Consolidated pivot report created: {output_path}")
            else:
                print(f"No data consolidated for {cono}.")

if __name__ == "__main__":
    consolidate_reports()
