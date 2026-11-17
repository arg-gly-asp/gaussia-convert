#!/usr/bin/env python
# coding: utf-8

# In[14]:

import argparse
from tester import TEST_MODEL_FILE
from openpyxl import load_workbook
import pandas as pd
import time
from datetime import datetime
import re
from collections import defaultdict
import os
from itertools import groupby

# #### Project: These set of functions together take an excel file that tracks a certain cell line growth in our lab and converts it into a dataframe format.

# In[15]:


### Step 1: Classify sheet as 'good' format or 'bad' format

# In[16]:


### Step 2: For a good format sheet, On column b, look for the second date on that column. Look for uninterrupted chains of numbers where the chain length is divisble by 3 on that row with the date. 
##Once you've identified the chains, split that chain up. 
## If the chain is more than 1 cell from the date then don't proceed.
### Save the date, save the file name, get the passage from the sheet name. v1 you get from the first 3 in thatchain
### v2 you get from the second and v3 the third. The letter comes from counting the first multiple of the chain
#### Save that data in a dataframe.

# In[17]:


def load_workbooks(folder_path): #Loads all workbooks in a folder 
    workbooks = []
    for filename in os.listdir(folder_path):
        if filename.endswith(('.xlsx', '.xlsm')) and not filename.startswith('~$'):
            path = os.path.join(folder_path, filename)
            wb = load_workbook(path)
            workbooks.append({'filename': filename, 'path': path, 'workbook': wb})
    return workbooks

# In[18]:


def parse_gbx_string(s):
    match = re.match(r'^GBX(\d+)P(\d+)', s)
    if not match:
        raise ValueError(f"String '{s}' doesn't match expected GBX###P# format")
    
    model_number = match.group(1)
    passage_number = match.group(2)
    
    return model_number, passage_number

def get_model_number(s):
    return parse_gbx_string(s)[0]

def get_passage_number(s):
    return parse_gbx_string(s)[1]

def get_attempt_number(s):
    """GBX683P1_2 -> 2. No underscore (or nothing after it) -> 1."""
    return TEST_MODEL_FILE.parse_sheet_title(s)[2]

# In[19]:


def group_numbers(nums):
    trimmed = nums[:-4]
    result = {}
    for i, val in enumerate(trimmed):
        letter = chr(ord('A') + i // 3)
        result.setdefault(letter, []).append(val)

    # Drop any key whose group has a None, or doesn't have exactly 3 values
    result = {
        k: v for k, v in result.items()
        if len(v) == 3 and None not in v
    }
    return result

# In[20]:


def drop_non_numerics(lst):
    result = []
    for x in lst:
        if isinstance(x, bool):
            continue
        if isinstance(x, (int, float)):
            result.append(x)
        elif isinstance(x, str):
            try:
                result.append(float(x) if '.' in x else int(x))
            except ValueError:
                pass
    return result

# In[34]:


"""
QC pass for sgluc Excel files, meant to run BEFORE gaussia_values ingestion.

Strategy: skip + log. Nothing here modifies source files. Any sheet with a
flagged issue is left out of the "clean" list and written to the report,
so it can be fixed by hand in Excel and re-checked.

Checks:
- Sheet title parses into model_id/passage
- Every date row (col B) has a dividing None found via the same forward
  scan the ingestion script uses. A single row missing its divider isn't
  flagged; more than one missing divider in a sheet is treated as a
  real pattern and gets flagged.
- Value count before the divider is a clean multiple of 3 (triplets)

Assumes the same helper functions used in the ingestion script
(get_model_number, get_passage_number, load_workbooks) are importable
from your existing utils module — swap the import below to match.
"""

from datetime import datetime
import logging
from pathlib import Path

# from your_utils_module import load_workbooks, get_model_number, get_passage_number

logging.basicConfig(
    filename='qc_report.log',
    filemode='w',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)

DATA_START_COL = 3  # column C — first real-data column, matches ingestion script


def find_date_rows(ws):
    """Return list of row numbers in column B that hold a real datetime."""
    return [cell.row for cell in ws["B"] if isinstance(cell.value, datetime)]


def check_title_parses(ws_title):
    """Try the same title parsing the ingestion script relies on."""
    try:
        model_id = get_model_number(ws_title)
        passage = get_passage_number(ws_title)
        if model_id is None or passage is None:
            return None, None, f"Title '{ws_title}' parsed but returned None"
        return model_id, passage, None
    except Exception as e:
        return None, None, f"Title '{ws_title}' failed to parse: {e}"


def scan_row(ws, row_num):
    """
    Replicates the ingestion script's forward scan from column C, looking
    for the "dividing None" (a None flanked by non-None cells on both
    sides). Returns (has_divider, values) where values are the real
    values collected before the divider (or before running off the end
    of the row, if no divider was found).
    """
    values = []
    has_divider = False
    for col_index in range(DATA_START_COL, ws.max_column + 1):
        target_cell = ws.cell(row=row_num, column=col_index)
        next_cell = ws.cell(row=row_num, column=col_index + 1)
        prev_cell = ws.cell(row=row_num, column=col_index - 1)
        if (target_cell.value is None
                and next_cell.value is not None
                and prev_cell.value is not None):
            has_divider = True
            break
        if(target_cell.value is None and next_cell.value is None):  #Checks if the next cells are empty
            n = 0
            found_non_empty = False
            while n < 9 and (col_index + n) <= ws.max_column:
                s_cell = ws.cell(row=row_num, column=col_index + n)
                if s_cell.value is not None:
                    found_non_empty = True
                    break
                n += 1
            if found_non_empty:
                continue
            has_divider = True
            break
        values.append(target_cell.value)
    return has_divider, values


def check_dividers_and_triplets(ws, date_rows):
    """
    For each date row, check whether a dividing None was found at all, and
    whether the values collected before it are a clean multiple of 3
    (group_numbers expects triplets).

    A single date row missing its divider isn't flagged on its own — but
    if MORE THAN ONE date row in the sheet is missing a divider, that's
    a pattern worth flagging rather than a one-off.
    """
    issues = []
    missing_divider_rows = []
    
    for row_num in date_rows:
        has_divider, values = scan_row(ws, row_num)
        #Have to drop all non-numeric values from values
        values = drop_non_numerics(values)
        
        if not has_divider:
            missing_divider_rows.append(row_num)
            continue
        if len(values) % 3 != 0:
            issues.append(
                f"Row {row_num}: {len(values)} data columns found before dividing gap, "
                f"not divisible by 3 (check for a missing/extra replicate or stray value)"
            )

    if len(missing_divider_rows) > 1:
        issues.append(
            f"{len(missing_divider_rows)} date row(s) with no dividing None found: "
            f"rows {missing_divider_rows}"
        )

    return issues


def qc_check_workbook(item):
    filename = item['filename']
    wb = item['workbook']
    result = {'filename': filename, 'clean_sheets': [], 'flagged_sheets': []}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        issues = []

        model_id, passage, title_err = check_title_parses(ws.title)
        if title_err:
            issues.append(title_err)

        date_rows = find_date_rows(ws)
        if not date_rows:
            issues.append("No datetime found in column B")
        else:
            issues.extend(check_dividers_and_triplets(ws, date_rows))

        if issues:
            for issue in issues:
                logging.warning(f"{filename} | {sheet_name} | {issue}")
            result['flagged_sheets'].append({'sheet': sheet_name, 'issues': issues})
        else:
            result['clean_sheets'].append(sheet_name)

    return result


# In[35]:


def run_qc(workbooks):
    """Run QC against an already-loaded list of workbook items (see
    load_workbooks). Returns the per-workbook results; does not load
    anything itself, so QC and ingestion always see the exact same
    in-memory sheets."""
    all_results = [qc_check_workbook(item) for item in workbooks]

    total_flagged = sum(len(r['flagged_sheets']) for r in all_results)
    total_clean = sum(len(r['clean_sheets']) for r in all_results)

    print(f"QC complete: {total_clean} sheet(s) clean, {total_flagged} sheet(s) flagged.")
    if total_flagged:
        print("\nFlagged sheets:")
        for r in all_results:
            for flagged in r['flagged_sheets']:
                print(f"\n  {r['filename']} | {flagged['sheet']}")
                for issue in flagged['issues']:
                    print(f"    - {issue}")
    return all_results


# In[36]:

def load_one(path):
    """Load a single workbook in the same item shape load_workbooks() returns."""
    return {'filename': os.path.basename(path), 'path': path, 'workbook': load_workbook(path)}


def designation_letters(ws, header_row):
    """Designation letters (A, B, C...) read from the header row that sits
    directly above the first date row. Reads left to right from column C and
    stops at the 'Negative' control column, so the control is never a letter."""
    letters = []
    for col in range(DATA_START_COL, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if isinstance(val, str):
            label = val.strip()
            if label.lower() == 'negative':
                break
            if len(label) == 1 and label.isalpha():
                letters.append(label.upper())
    return letters


def build_records(workbooks, sheet=None, designation=None):
    """Loop every sheet of every workbook. Returns (records, skipped, warnings).
    One TEST_MODEL_FILE per designation letter per sheet.

    The FIRST date row in column B is the injection date (day 0) and holds no
    measurements: it is recorded as first_date but never read for data. Real
    data starts on the second date row."""
    records, skipped, warnings = [], [], []
    for item in workbooks:
        wb = item['workbook']
        for sheet_name in wb.sheetnames:
            if sheet and sheet_name != sheet:
                continue
            ws = wb[sheet_name]
            try:
                TEST_MODEL_FILE.parse_sheet_title(ws.title)
            except ValueError as e:      # bad sheet title
                skipped.append((item['filename'], sheet_name, str(e)))
                continue
            date_rows = find_date_rows(ws)
            if not date_rows:
                skipped.append((item['filename'], sheet_name, "no datetime in column B"))
                continue
            if len(date_rows) < 2:
                skipped.append((item['filename'], sheet_name,
                                "only the first (day 0) date is present, no data rows yet"))
                continue
            letters = designation_letters(ws, date_rows[0] - 1)
            if not letters:
                skipped.append((item['filename'], sheet_name,
                                f"no designation letters found in header row {date_rows[0] - 1}"))
                continue
            # cross-check: triplets on the first data row = letters + 1 Negative control
            _, values = scan_row(ws, date_rows[1])
            n_triplets = len(drop_non_numerics(values)) // 3
            if n_triplets != len(letters) + 1:
                warnings.append((item['filename'], sheet_name,
                                 f"{len(letters)} letter(s) in header but {n_triplets} triplet(s) "
                                 f"on row {date_rows[1]} (expected letters + 1 Negative)"))
            for letter in letters:
                if designation and letter != designation.upper():
                    continue
                records.append(TEST_MODEL_FILE.from_worksheet(
                    item['filename'], ws, designation=letter))
    return records, skipped, warnings


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Build TEST_MODEL_FILE records from gaussia Excel files.")
    parser.add_argument("-d", help="Folder of workbooks to run.")
    parser.add_argument("-f", help="Single workbook file to run.")
    parser.add_argument("--sheet", help="Only process the sheet with this exact title, e.g. GBX683P1_2.")
    parser.add_argument("--designation", help="Only keep this designation letter, e.g. A.")
    args = parser.parse_args()
    try:
        if args.d:
            workbooks = load_workbooks(args.d)
        elif args.f:
            workbooks = [load_one(args.f)]
        else:
            workbooks = []
            print('No workbook folder or file selected.')
        if workbooks:
            records, skipped, warnings = build_records(workbooks, sheet=args.sheet,
                                                       designation=args.designation)
            df = pd.DataFrame([r.to_dict() for r in records])
            print(df.to_string(index=False) if not df.empty else 'No records built.')
            for fname, sname, reason in warnings:
                print(f'WARNING {fname} | {sname} | {reason}')
            for fname, sname, reason in skipped:
                print(f'SKIPPED {fname} | {sname} | {reason}')
    except Exception as e:
        print(str(e))
    finally:
        print('Done')