"""
FreightPOP Transaction Report Cleaner
=====================================
Daily LTL workflow automation:
  1. Deletes column A (Customer Id) and columns BA-BE
  2. Filters out blanks in column AL (Original Approved Invoice Amount)
  3. Filters out True in column AY (Exported to ERP)
  4. Groups rows by category, then sorts A→Z within each group:
        1) Audit customers      (Customer Audit = "On", minus overrides)
        2) Autopay customers    (configured list below)
        3) Priority Regular     (e.g. Associated Packaging Inc)
        4) Regular customers    (everything else)
  5. Writes formulas on every data row:
        AH = AL    (Rate without mark up)
        AI = AG-AH (Shipment Gross Profit)
        AN = AI/AH (Gross Profit %)
  6. Stamps today's date in column BA
  7. For Associated Packaging rows, copies the 'User' field (which holds
     values like 'Michael Fischer/AP100', 'Memphis AP400') into column BB
  8. Highlights every row where gross profit (AG − AL) is negative
  9. Outputs a cleaned .xlsx ready to paste into the master LTL file

Usage:
  python freightpop_cleaner.py <input_file.xlsx>
  python freightpop_cleaner.py            # auto-finds the newest .xlsx in folder
"""

import sys
import glob
import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# ---- Configuration ------------------------------------------------------
COLS_TO_DELETE = [
    'Customer Id',                    # original column A
    'User',                           # original BA
    'Appointment Set',                # original BB
    'Appointment Date',               # original BC
    'Parent Spot Quote Shipment Id',  # original BD
    'Spot Quote Fulfilled by',        # original BE
]

BLANK_FILTER_COL = 'Original Approved Invoice Amount'   # post-shift AL: drop blanks
TRUE_FILTER_COL  = 'Exported to ERP'                    # post-shift AY: drop True

# --- Customer category lists ---------------------------------------------
# All matching is CASE-INSENSITIVE and uses substring match, so partial
# names work (e.g. "Keystone" will catch "Keystone Safety Cleanroom Products").
# Edit any of these lists as customers move between groups.

# Customers in this list appear at the very TOP of the cleaned file.
# Source of truth: FreightPOP "Customer Audit = On" column, MINUS the
# AUDIT_OVERRIDES list below.
AUDIT_OVERRIDES = [           # treat these as Regular even if FreightPOP says Audit=On
    'Everflow Supplies LLC',
    'Keystone',
    'TileMart',
]

AUTOPAY_CUSTOMERS = [         # second tier — autopay customers
    'OptiMA Inc.',
    'Proline Range Hoods',
    'Global Auto Parts Inc',
    'Luxury 4 Less Appliances LLC',
    'Wirthco',
]

PRIORITY_REGULAR = [          # third tier — show at top of the Regular section
    'Associated Packaging',
]
# Everything else falls into the fourth tier: Regular (alphabetical).

# Formulas use the MASTER file's column letters. They'll show #NAME?/odd
# values inside the cleaned file but will calculate correctly once the
# rows are pasted into the master LTL workbook.
FORMULAS = {
    'AH': '=AL{row}',           # Rate without mark up
    'AI': '=AG{row}-AH{row}',   # Shipment Gross Profit
    'AN': '=AI{row}/AH{row}',   # Gross Profit %
}
DATE_COL = 'BA'                 # Cust Invoice Date — stamp today's date

# Highlight rows where Gross Profit (AG − AL) is negative.  Color is light red,
# the standard Excel "Bad" highlight.  Change the hex code below to use a
# different color (e.g. 'FFEB9C' for yellow, 'FFFF00' for bright yellow).
NEGATIVE_GP_FILL = PatternFill('solid', start_color='FFC7CE', end_color='FFC7CE')

# --- AP code extraction --------------------------------------------------
# For Associated Packaging rows only: the FreightPOP "User" column contains
# values like "Memphis AP400" or "Michael Fischer/AP100".  We paste the full
# value into column BB on each AP row.  Other customers get a blank BB.
AP_CUSTOMER_MATCH = 'Associated Packaging'
AP_SOURCE_COL     = 'User'
AP_TARGET_COL     = 'BB'
AP_HEADER         = 'AP Code'
# -------------------------------------------------------------------------


def _name_matches(customer_name: str, keyword_list: list[str]) -> bool:
    """True if customer_name contains any keyword (case-insensitive)."""
    nm = str(customer_name).lower().strip()
    return any(kw.lower().strip() in nm for kw in keyword_list if kw.strip())


def categorize(row) -> int:
    """0 = Audit, 1 = Autopay, 2 = Priority Regular, 3 = Regular."""
    name = str(row.get('Customer Name', ''))
    is_audit_flag = str(row.get('Customer Audit', '')).strip().lower() == 'on'

    # Audit, but only if NOT in the override list
    if is_audit_flag and not _name_matches(name, AUDIT_OVERRIDES):
        return 0
    if _name_matches(name, AUTOPAY_CUSTOMERS):
        return 1
    if _name_matches(name, PRIORITY_REGULAR):
        return 2
    return 3


def clean(input_path: str, output_path: str | None = None) -> str:
    df = pd.read_excel(input_path)
    rows_in = len(df)
    print(f"  Loaded:        {rows_in:>6,} rows × {df.shape[1]} cols")

    # 0) Capture AP code for Associated Packaging rows BEFORE we delete 'User'
    def _extract_ap(row):
        if AP_CUSTOMER_MATCH.lower() not in str(row.get('Customer Name', '')).lower():
            return ''
        val = row.get(AP_SOURCE_COL, '')
        return '' if pd.isna(val) else str(val).strip()
    df['_ap_code'] = df.apply(_extract_ap, axis=1)
    ap_found = (df['_ap_code'] != '').sum()
    print(f"  AP codes:      {ap_found:>6,} rows tagged for column {AP_TARGET_COL}")

    # 1) Delete columns
    missing = [c for c in COLS_TO_DELETE if c not in df.columns]
    if missing:
        print(f"  ⚠️  Headers not found, skipping: {missing}")
    df = df.drop(columns=[c for c in COLS_TO_DELETE if c in df.columns])
    print(f"  After delete:  {df.shape[1]} cols remain")

    # 2) Drop blanks in AL
    before = len(df)
    df = df[df[BLANK_FILTER_COL].notna()]
    print(f"  Blank filter:  removed {before - len(df):>5,} rows "
          f"(blank '{BLANK_FILTER_COL}')")

    # 3) Drop True in AY
    before = len(df)
    mask_true = df[TRUE_FILTER_COL].astype(str).str.strip().str.lower() == 'true'
    df = df[~mask_true]
    print(f"  True filter:   removed {before - len(df):>5,} rows "
          f"(True '{TRUE_FILTER_COL}')")

    # 4) Categorize + sort: Audit → Autopay → Regular, A→Z within each
    df['_category'] = df.apply(categorize, axis=1)
    df = (df.sort_values(by=['_category', 'Customer Name'],
                          key=lambda s: s.str.lower() if s.dtype == 'object' else s,
                          kind='stable')
            .reset_index(drop=True))
    cat_counts = df['_category'].value_counts().to_dict()
    print(f"  Categorized:   "
          f"{cat_counts.get(0,0)} Audit, "
          f"{cat_counts.get(1,0)} Autopay, "
          f"{cat_counts.get(2,0)} Priority Regular, "
          f"{cat_counts.get(3,0)} Regular")
    df = df.drop(columns=['_category'])

    # Capture AP codes aligned to the final (sorted) row order, then strip
    # the temp column so it doesn't appear in the output.
    ap_codes = df['_ap_code'].tolist()
    df = df.drop(columns=['_ap_code'])

    # Identify rows with negative gross profit (post-paste AN will be < 0).
    # Master formula: AI = AG - AH, where AG = Shipment Marked-Up Rate and
    # AH = AL = Original Approved Invoice Amount.  So we flag rows where
    # (Marked-Up Rate − Original Approved Invoice Amount) < 0.
    gp_diff = (df['Shipment Marked-Up Rate'].fillna(0)
               - df['Original Approved Invoice Amount'].fillna(0))
    negative_rows = [i for i, v in enumerate(gp_diff) if v < 0]
    print(f"  Negative GP:   {len(negative_rows):>5,} rows flagged for highlighting")

    # 5) Save data
    if output_path is None:
        stamp = datetime.now().strftime('%Y-%m-%d')
        base = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(os.path.dirname(input_path) or '.',
                                    f'{base}_CLEANED_{stamp}.xlsx')
    df.to_excel(output_path, index=False)

    # 6) Re-open with openpyxl to add formulas + date stamp
    wb = load_workbook(output_path)
    ws = wb.active
    today = datetime.now()
    last_row = ws.max_row  # includes header in row 1
    bold = Font(bold=True)

    # Overwrite BA header label
    ws[f'{DATE_COL}1'] = 'Cust Invoice Date'
    ws[f'{DATE_COL}1'].font = bold

    # Add BB header for AP codes
    ws[f'{AP_TARGET_COL}1'] = AP_HEADER
    ws[f'{AP_TARGET_COL}1'].font = bold

    for row in range(2, last_row + 1):
        for col, template in FORMULAS.items():
            ws[f'{col}{row}'] = template.format(row=row)
        ws[f'{DATE_COL}{row}'] = today
        ws[f'{DATE_COL}{row}'].number_format = 'm/d/yyyy'
        ws[f'AN{row}'].number_format = '0%'
        # AP code for Associated Packaging rows (blank for everyone else)
        ap_value = ap_codes[row - 2] if row - 2 < len(ap_codes) else ''
        if ap_value:
            ws[f'{AP_TARGET_COL}{row}'] = ap_value

    # Highlight every cell of each negative-GP row in light red
    for df_idx in negative_rows:
        excel_row = df_idx + 2     # df idx 0 → Excel row 2
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=excel_row, column=col_idx).fill = NEGATIVE_GP_FILL

    wb.save(output_path)

    print(f"\n✅  Output:      {len(df):>6,} rows  →  {output_path}")
    print(f"   Sort order:  Audit → Autopay → Priority Regular → Regular (A→Z within each)")
    print(f"   Formulas:    AH = AL,  AI = AG-AH,  AN = AI/AH (master letters)")
    print(f"   Date stamp:  {DATE_COL} = {today.strftime('%m/%d/%Y')}")
    print(f"   Net change:  {rows_in:,} → {len(df):,} rows ({rows_in - len(df):,} removed)")
    print(f"\n   Note:  formulas will recalculate correctly once pasted into")
    print(f"          the master LTL file.")
    return output_path


def find_latest_xlsx(folder: str = '.') -> str:
    candidates = [f for f in glob.glob(os.path.join(folder, '*.xlsx'))
                  if 'CLEANED' not in f and not os.path.basename(f).startswith('~$')]
    if not candidates:
        sys.exit("❌  No .xlsx files found in current folder.")
    return max(candidates, key=os.path.getmtime)


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else find_latest_xlsx()
    print(f"📂 Processing:  {src}\n")
    clean(src)
