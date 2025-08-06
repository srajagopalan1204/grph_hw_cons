# Multi-Cono Consolidation (Pivot by Date_of_rep)

## Overview
- Merges section files across multiple Date_of_rep values into a single sheet per section.
- Column headers format: `<MMDDYYYY>_<CompCol>`
- Uses key columns defined in config.json5.

## Setup
1. Upload this project ZIP to Codespaces
2. Upload `Cono1_Weekly_reports.zip` and `Cono3_Weekly_reports.zip`
3. Unzip:
```bash
mkdir -p Cono1/Weekly_reports Cono3/Weekly_reports
unzip Cono1_Weekly_reports.zip -d Cono1/Weekly_reports
unzip Cono3_Weekly_reports.zip -d Cono3/Weekly_reports
```
4. Install dependencies:
```bash
pip install -r requirements.txt
```
5. Run consolidation:
```bash
python consolidate_reports.py
```
6. Download consolidated workbooks from `ConoX/Consolidated_reports` folders.
