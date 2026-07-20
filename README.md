# oem-rate-sheet-automation
Python tool for automated OEM finance rate sheet extraction, normalization, and change detection
# OEM Finance Rate Sheet Automation
**Python | pandas | pdfplumber | openpyxl | Excel**

## The Problem
Every finance cycle, vehicle manufacturers (OEMs) send updated consumer 
lending rates — each in a completely different format. Some deliver Excel 
tables, some matrices, some messy merged-cell layouts, some PDF. Comparing 
each sheet against the last cycle by hand was slow and error-prone.

## What This Tool Does
1. Reads each OEM's rate sheet using a dedicated parser per format
2. Normalizes everything into one clean structure (OEM, Model, Term, Rate)
3. Compares against the prior cycle to detect what changed, what's new, 
   and what was dropped
4. Outputs a structured Excel workbook with four tabs:
   - **Clean_Data** — standardized rates ready for entry
   - **Changes** — rates that moved, highlighted in red
   - **Added_Vehicles** — new models this cycle
   - **Removed_Vehicles** — models dropped this cycle

## Result
Cut a multi-hour manual review down to under two minutes.  
Originally prototyped in Excel VBA, then rebuilt in Python for speed.

## Files
- `oem_rate_sheet_automation.py` — full Python source code
- `oem_rate_sheet_automation.pdf` — portfolio writeup with output screenshots
- `oem_rate_sheet_clean_output.xlsx` — sample Excel output

## Note
Demo files use fictional OEM names and fabricated data for portfolio 
purposes. Built at EasyDeal–AutoSync (2025).

Author: Majd Kabbani | [LinkedIn](https://www.linkedin.com/in/majd-kabbani-39335892)
