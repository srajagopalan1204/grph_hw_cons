# Graph Generation Module (GitHub Codespaces)

This module creates Excel line charts for consolidated reports for each Cono (e.g., Cono1, Cono2, Cono3). It automatically identifies the latest report per Cono and generates charts based on the configuration.

## 📦 Initial Setup Steps (after cloning the base repository)

1. **Upload the ZIP Package**
   - Upload `graph_dynamic_package.zip` to your Codespaces environment.

2. **Unzip the Package**
   ```bash
   unzip graph_dynamic_package.zip
   cd graph_dynamic_package
   ```

3. **Install Required Packages**
   ```bash
   pip install -r requirements.txt
   ```

4. **Prepare Report Files**
   - Ensure that the folder structure exists:
     ```
     ./Cono1/Consolidated_reports/
     ./Cono2/Consolidated_reports/
     ./Cono3/Consolidated_reports/
     ```
   - Place your Excel report(s) for each Cono into their respective folders.
   - The script will automatically detect the latest `.xlsx` file in each.

---

## 📈 How to Create the Report

1. **Configure the Charts in `config_grph.json`**
   ```json
   {
     "workbooks": [
       {
         "cono": "Cono1",
         "charts": [
           {
             "sheet": "Section1_TakenBy_OEcount",
             "chart_type": "line",
             "title": "OE Count Trend by TakenBy",
             "x_col": "TakenBy",
             "y_col": ["08052025_OE_Count", "08042025_OE_Count"],
             "output_sheet": "Graph_TakenBy_OEcount"
           }
         ]
       },
       {
         "cono": "Cono2",
         "charts": [ ... ]
       },
       {
         "cono": "Cono3",
         "charts": [ ... ]
       }
     ]
   }
   ```

2. **Run the Script**
   ```bash
   python generate_graphs.py
   ```

3. **Output**
   - A new file will be created for each Cono in the format:
     ```
     Cono2_Consolidate_report_XXXX_graph_YYYYMMDD_HHMM.xlsx
     ```
   - Chart tabs will be inserted at the beginning of the file.

---

## 🛠 Notes
- Make sure each `workbook` entry in the config includes a `"cono"` key with a valid folder name (`Cono1`, `Cono2`, `Cono3`).
- Invalid or missing `cono` entries will be skipped with a warning.
