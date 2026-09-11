import argparse
import csv
import os
import sys
import re
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ==========================================
# USER CONFIGURATION
# ==========================================

# 1. Define your 5-minute grace period for "on-time" calculations.
ON_TIME_TOLERANCE_MINUTES = 5

def _config_value(secret_key, default):
    """
    Reads a config value from Streamlit secrets when available (the
    hosted app's repo is public, so real business identifiers like the
    broker name live only in Streamlit Cloud's Secrets manager, not in
    this source file), otherwise falls back to `default` below - which
    is what local/CLI use without any secrets.toml will get.
    """
    try:
        import streamlit as st
        val = st.secrets.get(secret_key)
        if val:
            return val
    except Exception:
        pass
    return default


# 2. Fixed values that go on every row (edit these to match your account -
# or, for the hosted app, set supplier_name / contract_id in Secrets so
# they don't need to sit in this public source file).
SUPPLIER_NAME = _config_value('supplier_name', 'Your Broker Name')
CONTRACT_ID = _config_value('contract_id', 'YOUR-CONTRACT-ID')

# Shown as a signature line at the top of the Overview tab.
PREPARED_BY = 'Lamarr Glover'

# 3. Where your Google OAuth credentials.json lives. Defaults to the same
#    folder this script file is in, so just keep credentials.json next to
#    build_trip_report.py and you don't need to touch this.
GOOGLE_CREDENTIALS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'credentials.json'
)


# ==========================================
# COLUMN MAPPING (by header NAME, not position)
# ==========================================
#
# The script doesn't rely on a column ever being "the Nth column" - it
# looks up each field by its header text in row 1. That means you can
# add, remove, or reorder columns in the sheet and this keeps working, as
# long as the header names below still match what's in row 1.
#
# Left side = internal field name (used elsewhere in this script).
# Right side = the exact header text to look for (whitespace/case don't
# matter - "PU Arr Reason" matches "pu arr reason" or a wrapped two-line
# header just fine).
FIELD_HEADERS = {
    'trip_id': '119498370',  # sheet's actual header for the trip column
    'load_id': 'Load',
    'origin': 'ORIGIN',
    'pu_date': 'PU Date',
    'pu_status': 'PU STATUS',
    'pu_arrival': 'PU Arrival',
    'pu_arr_reason': 'PU Arr Reason',
    'pu_departure': 'PU Departure',
    'pu_dp_reason': 'PU DP Reason',
    'dest': 'DEST',
    'del_date': 'DEL Date',
    'del_status': 'DEL STATUS',
    'del_arrival': 'DEL Arrival',
    'del_arr_reason': 'DEL Arr Reason',
    'del_departure': 'DEL Departure',
    'driver_name': 'Driver Name',
    'driver_phone': 'Driver Phone',
    'carrier': 'Carrier',
    'tracking': 'Tracking',
    'notes': 'Notes',
    'dispatcher': 'Dispatcher',
    'lane': 'Lane department OD pair',
    'spot_origin': 'Spot Origin',
    'spot_dest': 'Spot Dest',
    'report_pu': 'Report PU',
    'report_del': 'Report DEL',
    'created_date': 'Created date',
    'miles': 'Miles',
}

# Fields that must be present for the script to run at all. Everything
# else in FIELD_HEADERS is optional - missing ones just come back blank.
# trip_id is included here so a header mismatch fails loudly instead of
# silently leaving every row's trip_id blank (which is what caused the
# --trips filter to only match on Load ID before).
REQUIRED_FIELDS = [
    'trip_id', 'load_id', 'origin', 'pu_date', 'pu_arrival', 'pu_departure',
    'dest', 'del_date', 'del_arrival', 'del_status', 'lane',
]

# Any PU/DEL status containing one of these substrings means the load
# never happened, so it's excluded from the report entirely.
EXCLUDE_STATUS_KEYWORDS = ('TONU', 'CANCEL')


def normalize_header(text):
    """Collapse whitespace/newlines and lowercase, for tolerant matching."""
    return re.sub(r'\s+', ' ', (text or '').strip()).lower()


def build_field_index(header_row):
    """
    Map internal field name -> column index, by matching header_row
    against FIELD_HEADERS. Raises a clear error naming any missing
    required header, instead of silently misreading columns.
    """
    normalized_to_idx = {
        normalize_header(cell): idx for idx, cell in enumerate(header_row)
    }

    field_index = {}
    missing = []

    for field, expected_header in FIELD_HEADERS.items():
        idx = normalized_to_idx.get(normalize_header(expected_header))
        if idx is None:
            if field in REQUIRED_FIELDS:
                missing.append(expected_header)
        else:
            field_index[field] = idx

    if missing:
        raise ValueError(
            "Could not find these required columns in row 1: "
            + ", ".join(missing)
            + ". Check that the header text in your sheet matches "
              "FIELD_HEADERS at the top of this script."
        )

    return field_index


def parse_datetime(dt_str):
    """Parser: cleans dirty strings and tries multiple formats."""
    if not dt_str or not isinstance(dt_str, str):
        return None

    # Some cells accumulate more than one timestamp stacked on separate
    # lines - e.g. an appointment that got rescheduled, with the newer
    # time appended below the original ("07/05/26 17:35\n07/05/26
    # 19:00"). Collapsing that straight to one space (like below) would
    # glue both dates into one unparseable string, so instead take the
    # last non-blank line - the most recently entered value - before any
    # further cleanup.
    lines = [ln.strip() for ln in dt_str.splitlines() if ln.strip()]
    if not lines:
        return None

    cleaned = lines[-1]
    cleaned = cleaned.replace(' :', ' ')
    cleaned = re.sub(r'\s+', ' ', cleaned)

    formats = [
        '%m/%d/%y %H:%M',
        '%m/%d/%Y %H:%M',
        '%m/%d/%y %H:%M:%S',
        '%m/%d/%Y %H:%M:%S',
    ]

    for fmt in formats:
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue

    return None


def is_on_time(scheduled, actual):
    if scheduled is None or actual is None:
        return 'NO'

    if actual <= scheduled + timedelta(minutes=ON_TIME_TOLERANCE_MINUTES):
        return 'YES'

    return 'NO'


def is_excluded_status(status):
    """True if a PU/DEL status marks the load as cancelled or TONU."""
    s = (status or '').upper()
    return any(keyword in s for keyword in EXCLUDE_STATUS_KEYWORDS)


def parse_id_list(text):
    """
    Split a user-supplied list of trip/load numbers on commas and/or
    whitespace, so it works whether it's typed as "53, 55, 121744191" or
    pasted as a whole column copied straight out of Excel (one number per
    line). Returns None (meaning "no filter") for empty input.
    """
    if not text:
        return None
    ids = {t.strip() for t in re.split(r'[,\s]+', text) if t.strip()}
    return ids or None


def clean_lane(lane_str):
    """
    The sheet already gives us a ready-made "Origin - Dest" lane string
    (the "Lane department OD pair" column), so we just clean it up rather
    than reconstruct it - except for a same-city self-loop, which we
    still normalize to "n/a".
    """
    if not lane_str:
        return 'n/a'

    normalized = re.sub(r'[\u2013\u2014]', '-', lane_str).strip()
    parts = [p.strip() for p in normalized.split(' - ') if p.strip()]

    if len(parts) >= 2 and parts[0] == parts[-1]:
        return 'n/a'

    return normalized if normalized else 'n/a'


def build_notes(pu_arr_reason, pu_dp_reason, del_arr_reason, general_notes,
                 on_time_arr, on_time_dispatch, on_time_delivery):
    """
    Auto-fill the "REQUIRED IF PICK UP OR DELIVERY IS MARKED NO" column
    from the reason fields the sheet already tracks, instead of leaving
    it for manual typing.
    """
    parts = []

    if on_time_arr == 'NO' and pu_arr_reason:
        parts.append(f"Pickup late: {pu_arr_reason}")

    if on_time_dispatch == 'NO' and pu_dp_reason:
        parts.append(f"Dispatch late: {pu_dp_reason}")

    if on_time_delivery == 'NO' and del_arr_reason:
        parts.append(f"Delivery late: {del_arr_reason}")

    if general_notes:
        parts.append(general_notes)

    return ' | '.join(parts)


def parse_rows(rows, start_date=None, end_date=None, trip_filter=None,
                carrier_filter=None):
    """
    Shared parsing logic. `rows` is a list of row-lists (list of string
    cell values), with row[0] being the header row - works the same
    whether it came from a local tab-delimited file or a Google Sheet.

    start_date / end_date: date objects (inclusive) used to filter on the
    load's PU Date. Either or both can be None to leave that side open.

    trip_filter: optional set of strings. If given, a row is only kept
    when its Trip # OR its Load # matches something in the set - so you
    can pass either kind of number and it just works.

    carrier_filter: optional set of carrier-name strings (matched
    case-insensitively, exact match against the Carrier column). If
    given, a row is only kept when its carrier is in the set.
    """
    if not rows:
        return [], {}

    field_index = build_field_index(rows[0])

    def get(row, field):
        idx = field_index.get(field)
        if idx is None or idx >= len(row):
            return ''
        return (row[idx] or '').strip()

    carrier_filter_upper = (
        {c.upper() for c in carrier_filter} if carrier_filter else None
    )

    loads = []
    stats = {
        'total_rows': 0,
        'skipped_blank': 0,
        'skipped_cancelled_tonu': 0,
        'skipped_date_range': 0,
        'skipped_unparseable_date': 0,
        'skipped_trip_filter': 0,
        'skipped_carrier_filter': 0,
        'skipped_missing_checkin': 0,
        'included': 0,
    }

    for row in rows[1:]:
        if not row or not any(c.strip() for c in row if c):
            stats['skipped_blank'] += 1
            continue

        stats['total_rows'] += 1

        trip_id = get(row, 'trip_id')
        load_id = get(row, 'load_id')
        carrier = get(row, 'carrier')
        pu_status = get(row, 'pu_status')
        del_status = get(row, 'del_status')

        if is_excluded_status(pu_status) or is_excluded_status(del_status):
            stats['skipped_cancelled_tonu'] += 1
            continue

        if carrier_filter_upper is not None and carrier.upper() not in carrier_filter_upper:
            stats['skipped_carrier_filter'] += 1
            continue

        dt_pu_date = parse_datetime(get(row, 'pu_date'))

        if start_date is not None or end_date is not None:
            if dt_pu_date is None:
                stats['skipped_unparseable_date'] += 1
                continue
            pu_day = dt_pu_date.date()
            if start_date is not None and pu_day < start_date:
                stats['skipped_date_range'] += 1
                continue
            if end_date is not None and pu_day > end_date:
                stats['skipped_date_range'] += 1
                continue

        if trip_filter is not None:
            if trip_id not in trip_filter and load_id not in trip_filter:
                stats['skipped_trip_filter'] += 1
                continue

        lane = clean_lane(get(row, 'lane'))

        dt_pu_arrival = parse_datetime(get(row, 'pu_arrival'))
        dt_pu_departure = parse_datetime(get(row, 'pu_departure'))
        dt_del_date = parse_datetime(get(row, 'del_date'))
        dt_del_arrival = parse_datetime(get(row, 'del_arrival'))

        # Can't evaluate on-time performance without an actual check-in/
        # check-out, so a trip missing any of these is excluded entirely
        # rather than being counted as a late "NO".
        if dt_pu_arrival is None or dt_pu_departure is None or dt_del_arrival is None:
            stats['skipped_missing_checkin'] += 1
            continue

        dt_planned_dispatch = (
            dt_pu_date + timedelta(minutes=30) if dt_pu_date else None
        )

        on_time_arrival = is_on_time(dt_pu_date, dt_pu_arrival)
        on_time_dispatch = is_on_time(dt_planned_dispatch, dt_pu_departure)
        on_time_delivery = is_on_time(dt_del_date, dt_del_arrival)

        notes = build_notes(
            get(row, 'pu_arr_reason'),
            get(row, 'pu_dp_reason'),
            get(row, 'del_arr_reason'),
            get(row, 'notes'),
            on_time_arrival,
            on_time_dispatch,
            on_time_delivery,
        )

        loads.append({
            'load_id': load_id,
            'trip_id': trip_id,
            'carrier': carrier or 'Unknown Carrier',
            'lane': lane,
            'sched_arr': dt_pu_date,
            'actual_arr': dt_pu_arrival,
            'on_time_arr': on_time_arrival,
            'actual_dispatch': dt_pu_departure,
            'planned_dispatch': dt_planned_dispatch,
            'on_time_dispatch': on_time_dispatch,
            'actual_delivery': dt_del_arrival,
            'planned_delivery': dt_del_date,
            'on_time_delivery': on_time_delivery,
            'notes': notes,
        })
        stats['included'] += 1

    return loads, stats


def load_trips(input_file, start_date=None, end_date=None, trip_filter=None,
                carrier_filter=None):
    """Read trips from a local tab-delimited file. Row 1 must be headers."""
    with open(input_file, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='\t', quoting=csv.QUOTE_MINIMAL)
        rows = list(reader)

    return parse_rows(rows, start_date=start_date, end_date=end_date,
                       trip_filter=trip_filter, carrier_filter=carrier_filter)


def get_gspread_client():
    """
    Returns an authorized gspread client. Tries, in order:

    1. A Google service account defined in Streamlit secrets
       (st.secrets["gcp_service_account"]). This is what the hosted
       (Streamlit Community Cloud) app uses - a service account logs in
       with its own key file, no interactive browser consent needed,
       which matters because a headless host has no browser to pop open
       for the OAuth flow below anyway. Every visitor to the hosted app
       shares this same read-only identity; access is controlled by
       which sheets you've shared with the service account's email, not
       by who's asking.
    2. The interactive OAuth flow (gspread.oauth()) used for local/CLI
       use - opens a browser for a one-time login tied to your own
       Google account, then caches the result.

    See the "GOOGLE SHEETS SETUP" and "HOSTING SETUP" comment blocks
    below for the one-time setup steps for each.
    """
    try:
        import gspread
    except ImportError:
        raise RuntimeError(
            "gspread is not installed. Run:\n"
            "    pip install gspread google-auth-oauthlib"
        )

    try:
        import streamlit as st
        if 'gcp_service_account' in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets['gcp_service_account']))
    except ImportError:
        pass  # streamlit isn't installed (e.g. plain CLI use) - fall through to OAuth.
    except Exception:
        pass  # no secrets.toml / no such key - fall through to OAuth.

    # Opens a browser for a one-time login the first time it's run, then
    # reuses the cached credentials on subsequent runs.
    if GOOGLE_CREDENTIALS_PATH:
        return gspread.oauth(credentials_filename=GOOGLE_CREDENTIALS_PATH)
    return gspread.oauth()


def fetch_sheet_rows(sheet_url_or_key, worksheet_name=None):
    """
    Read the raw rows (list of list-of-strings, row 1 = headers) from a
    Google Sheet. Split out from load_trips_from_sheet() so a caller
    (e.g. the Streamlit app) can fetch once and then re-run parse_rows()
    against the cached rows for every filter change, instead of hitting
    Google Sheets again each time.
    """
    gc = get_gspread_client()

    if sheet_url_or_key.startswith('http'):
        sheet = gc.open_by_url(sheet_url_or_key)
    else:
        sheet = gc.open_by_key(sheet_url_or_key)

    ws = sheet.worksheet(worksheet_name) if worksheet_name else sheet.sheet1

    return ws.get_all_values()


def load_trips_from_sheet(sheet_url_or_key, worksheet_name=None,
                           start_date=None, end_date=None, trip_filter=None,
                           carrier_filter=None):
    """Read + filter trips directly from a Google Sheet in one call (CLI path)."""
    try:
        rows = fetch_sheet_rows(sheet_url_or_key, worksheet_name)
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    return parse_rows(rows, start_date=start_date, end_date=end_date,
                       trip_filter=trip_filter, carrier_filter=carrier_filter)


# ==========================================
# GOOGLE SHEETS SETUP (one-time)
# ==========================================
#
# 1. pip install gspread google-auth-oauthlib
# 2. In Google Cloud Console: create a project, enable the "Google Sheets
#    API", create an OAuth client ID of type "Desktop app", download the
#    JSON and save it as credentials.json next to this script.
# 3. In "Google Auth Platform" > "Audience", add your Gmail as a test
#    user (needed while the app is in Testing mode).
# 4. Run the script once with --sheet. A browser window opens asking you
#    to log in with the Google account that has access to the sheet
#    (gspread caches the login afterward, so you won't be asked again).
#
# This is only for running locally/CLI - it needs a real browser to pop
# open, so it can't run on a hosted server. See HOSTING SETUP below for
# the app.py / Streamlit Community Cloud path.


# ==========================================
# HOSTING SETUP - Google service account (one-time, for the hosted app)
# ==========================================
#
# The interactive OAuth login above can't run on a headless host (no
# browser to open), so the deployed app instead logs in as a Google
# "service account" - a robot account with its own key, no human login
# step. See README.md for the full walkthrough; short version:
#
# 1. In Google Cloud Console: same project as above (or a new one) ->
#    IAM & Admin > Service Accounts > Create Service Account. Any
#    name is fine.
# 2. Open that service account > Keys > Add Key > Create new key > JSON.
#    This downloads a .json key file - keep it private, never commit it
#    to GitHub.
# 3. Open your Google Sheet > Share > paste the service account's email
#    (looks like ...@...iam.gserviceaccount.com, found in the JSON key
#    or the service account's details page) > give it Viewer access.
# 4. On Streamlit Community Cloud, open your app > Settings > Secrets,
#    and paste the entire JSON key's contents under a [gcp_service_account]
#    section (see .streamlit/secrets.toml.example in this repo for the
#    exact format), plus an app_password entry for the login gate below.
#
# Locally, you can test this same path by copying
# .streamlit/secrets.toml.example to .streamlit/secrets.toml and filling
# it in - that file is gitignored, so it's safe to put real secrets there.


# ==========================================
# EXCEL OUTPUT
# ==========================================

# Headers exactly as they appear in the reference report (row 1, merged
# down into row 2 for columns A-O; P/Q/R are single-row labels above the
# live formula in row 2).
HEADERS = [
    'SUPPLIER NAME', 'Contract ID', 'SV Trip ID', 'Load ID', 'Carrier',
    'O/D PAIR',
    'Scheduled arrival time', 'Actual arrival time', 'ON TIME Arrival Y/N',
    'actual dispatch time', 'planned dispatch time', 'Dispatch on time Y/N',
    'Actual delivery time', 'planned delivery time', 'ON TIME DELIVERY y/n',
    'NOTES REQUIRED IF PICK UP OR DELIVERY IS MARKED "NO"',
]

SUMMARY_HEADERS = ['OT Arrival %', 'OT Dispatch %', 'OT Delivery %']

# 1-indexed columns holding datetimes / Yes-No values within HEADERS.
# (Shifted by 1 vs. the original reference layout to make room for the
# Carrier column inserted at position 5.)
DATE_COLS = {7, 8, 10, 11, 13, 14}
YESNO_COLS = {9, 12, 15}

# Which on-time field (from the `loads` dicts) each summary percentage is
# built from.
SUMMARY_SOURCE_FIELD = {
    'OT Arrival %': 'on_time_arr',
    'OT Dispatch %': 'on_time_dispatch',
    'OT Delivery %': 'on_time_delivery',
}

FONT = Font(name='Arial', size=10)
BOLD_FONT = Font(name='Arial', size=10, bold=True)
THIN = Side(style='thin', color='000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Matches the column widths measured from the reference file, with a
# width for the inserted Carrier column (position 5) added in.
COLUMN_WIDTHS = [
    19.86, 11.57, 7.71, 12.0, 18.0, 40.86, 19.71, 17.0, 6.29,
    17.57, 19.43, 6.57, 18.86, 19.86, 6.29, 41.0, 9.14, 9.14, 9.14,
]

INVALID_SHEET_CHARS = re.compile(r'[\\/*?:\[\]]')


def sanitize_sheet_name(name_source, used_names):
    """
    Excel sheet names: <=31 chars, no \\ / * ? : [ ], not blank,
    and must be unique within the workbook. Works the same whether
    name_source is a lane string or a trip ID.
    """
    name = INVALID_SHEET_CHARS.sub('', name_source).strip() or 'Sheet'
    name = name[:31]

    base = name
    suffix = 2
    while name.lower() in used_names:
        cut = 31 - len(f' ({suffix})')
        name = f'{base[:cut]} ({suffix})'
        suffix += 1

    used_names.add(name.lower())
    return name


def compute_on_time_pct(group_loads, field):
    """
    Percentage of loads in this group where `field` is 'YES', out of
    loads where that field was actually determinable (YES or NO both
    count - only missing/undetermined data is excluded). Returns a plain
    float (0-1) instead of an Excel formula so the number is correct the
    instant the file is opened, in any viewer - no recalculation pass
    needed.

    Returns None when there's nothing to measure (avoids a ZeroDivision
    and avoids reporting a misleading 0%).
    """
    values = [ld[field] for ld in group_loads if ld[field] in ('YES', 'NO')]
    if not values:
        return None
    yes_count = sum(1 for v in values if v == 'YES')
    return yes_count / len(values)


def write_lane_sheet(ws, loads):
    # --- Header: row 1 merged into row 2 for columns A-O ---
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws.cell(row=2, column=col_idx).border = BORDER
        ws.merge_cells(
            start_row=1, start_column=col_idx, end_row=2, end_column=col_idx
        )

    # --- Summary headers + computed percentages (P/Q/R) ---
    # Computed directly in Python from the same on-time YES/NO values
    # used in the data rows, and written as plain numbers (not formulas),
    # so they always display correctly - see compute_on_time_pct().
    summary_start_col = len(HEADERS) + 1  # column P

    for i, label in enumerate(SUMMARY_HEADERS):
        col_idx = summary_start_col + i
        header_cell = ws.cell(row=1, column=col_idx, value=label)
        header_cell.font = FONT
        header_cell.alignment = CENTER
        header_cell.border = BORDER

        pct = compute_on_time_pct(loads, SUMMARY_SOURCE_FIELD[label])

        pct_cell = ws.cell(row=2, column=col_idx, value=pct)
        pct_cell.number_format = '0%'
        pct_cell.font = BOLD_FONT
        pct_cell.alignment = CENTER
        pct_cell.border = BORDER

    # --- Data rows ---
    first_data_row = 3
    for r, load in enumerate(loads, start=first_data_row):
        row_values = [
            SUPPLIER_NAME,
            CONTRACT_ID,
            load['trip_id'],
            load['load_id'],
            load['carrier'],
            load['lane'],
            load['sched_arr'],
            load['actual_arr'],
            load['on_time_arr'],
            load['actual_dispatch'],
            load['planned_dispatch'],
            load['on_time_dispatch'],
            load['actual_delivery'],
            load['planned_delivery'],
            load['on_time_delivery'],
            load['notes'],
        ]

        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=r, column=col_idx, value=value)
            cell.font = FONT
            cell.border = BORDER
            cell.alignment = CENTER

            if col_idx in DATE_COLS:
                cell.number_format = 'm/d/yy h:mm'

    # --- Column widths / freeze / row heights ---
    for col_idx, width in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = 'A3'
    ws.row_dimensions[1].height = 30


def _write_summary_table(ws, start_row, label_header, groups):
    """
    Writes one [label | Loads | OT Arrival % | OT Dispatch % | OT Delivery %]
    table starting at start_row (a bold header row), one row per group key
    plus a bold TOTAL row at the bottom. Returns the row index just past
    the table (i.e. where a following table/title can start).
    """
    headers = [label_header, 'Loads', 'OT Arrival %', 'OT Dispatch %', 'OT Delivery %']

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.font = BOLD_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    r = start_row + 1
    for key, group_loads in groups.items():
        row_values = [
            key,
            len(group_loads),
            compute_on_time_pct(group_loads, 'on_time_arr'),
            compute_on_time_pct(group_loads, 'on_time_dispatch'),
            compute_on_time_pct(group_loads, 'on_time_delivery'),
        ]
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=r, column=col_idx, value=value)
            cell.font = FONT
            cell.border = BORDER
            cell.alignment = CENTER
            if col_idx >= 3:
                cell.number_format = '0%'
        r += 1

    # Overall total row across every load in this table.
    all_loads = [ld for group_loads in groups.values() for ld in group_loads]
    total_row = [
        'TOTAL',
        len(all_loads),
        compute_on_time_pct(all_loads, 'on_time_arr'),
        compute_on_time_pct(all_loads, 'on_time_dispatch'),
        compute_on_time_pct(all_loads, 'on_time_delivery'),
    ]
    for col_idx, value in enumerate(total_row, start=1):
        cell = ws.cell(row=r, column=col_idx, value=value)
        cell.font = BOLD_FONT
        cell.border = BORDER
        cell.alignment = CENTER
        if col_idx >= 3:
            cell.number_format = '0%'

    return r + 1


def write_overview_sheet(ws, groups, group_by, carrier_groups=None, prepared_by=None):
    """
    One-page summary across every lane/trip tab: row per group with its
    load count and OT Arrival/Dispatch/Delivery %, plus an overall total
    row at the bottom. Lets you see fleet-wide performance without
    opening every tab.

    When carrier_groups is given, a second table breaking the same
    metrics down by Carrier is written below the first, with a blank row
    and title in between.

    prepared_by: optional name shown as a signature line above the
    tables, with the generation date/time.
    """
    title_cell = ws.cell(row=1, column=1, value='TRIP ON-TIME PERFORMANCE REPORT')
    title_cell.font = Font(name='Arial', size=14, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

    generated = datetime.now().strftime('%m/%d/%Y %H:%M')
    subtitle = f'Generated: {generated}'
    if prepared_by:
        subtitle = f'Prepared by: {prepared_by}   |   {subtitle}'
    subtitle_cell = ws.cell(row=2, column=1, value=subtitle)
    subtitle_cell.font = Font(name='Arial', size=10, italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    table_start_row = 4
    label_header = 'Trip ID' if group_by == 'trip' else 'O/D Lane'
    next_row = _write_summary_table(ws, table_start_row, label_header, groups)

    if carrier_groups:
        title_row = next_row + 1
        carrier_title_cell = ws.cell(row=title_row, column=1, value='By Carrier')
        carrier_title_cell.font = BOLD_FONT
        _write_summary_table(ws, title_row + 1, 'Carrier', carrier_groups)

    ws.column_dimensions['A'].width = 40
    for col in ('B', 'C', 'D', 'E'):
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = f'A{table_start_row + 1}'


def write_excel(loads, output_file, group_by='lane'):
    """
    group_by: 'lane' (default) groups sheets by O/D lane, same as before.
    'trip' groups sheets by trip_id instead - one tab per trip number -
    which is what you want when you've filtered down to a handful of
    specific trips with --trips and want each on its own page.
    """
    wb = Workbook()
    wb.remove(wb.active)

    if group_by == 'trip':
        group_field = 'trip_id'
        fallback_label = 'No Trip ID'
    else:
        group_field = 'lane'
        fallback_label = 'Lane'

    # Group loads by the chosen field, preserving first-seen order.
    groups = {}
    for load in loads:
        key = load[group_field] or fallback_label
        groups.setdefault(key, []).append(load)

    # Carrier breakdown for the Overview tab - grouped independently of
    # the per-sheet grouping above, so it shows up whether the tabs
    # themselves are split by lane or by trip.
    carrier_groups = {}
    for load in loads:
        carrier_groups.setdefault(load['carrier'], []).append(load)

    # Overview tab first, so it's the first thing you see when you open
    # the workbook.
    overview_ws = wb.create_sheet(title='Overview')
    write_overview_sheet(overview_ws, groups, group_by, carrier_groups, prepared_by=PREPARED_BY)

    used_names = {'overview'}
    for key, group_loads in groups.items():
        sheet_name = sanitize_sheet_name(key, used_names)
        ws = wb.create_sheet(title=sheet_name)
        write_lane_sheet(ws, group_loads)

    wb.active = 0
    wb.save(output_file)


def print_filter_summary(stats, start_date, end_date, trip_filter, carrier_filter=None):
    if start_date or end_date:
        rng = f"{start_date or 'earliest'} to {end_date or 'latest'}"
        print(f"Date filter (PU Date): {rng}")
    if trip_filter:
        print(f"Trip filter: {', '.join(sorted(trip_filter))}")
    if carrier_filter:
        print(f"Carrier filter: {', '.join(sorted(carrier_filter))}")

    print(
        f"Rows read: {stats['total_rows']} | "
        f"included: {stats['included']} | "
        f"skipped cancelled/TONU: {stats['skipped_cancelled_tonu']} | "
        f"skipped (outside date range): {stats['skipped_date_range']} | "
        f"skipped (unparseable date): {stats['skipped_unparseable_date']} | "
        f"skipped (trip filter): {stats['skipped_trip_filter']} | "
        f"skipped (carrier filter): {stats['skipped_carrier_filter']} | "
        f"skipped (no check-in/out): {stats['skipped_missing_checkin']}"
    )


def finish(loads, stats, output_file, start_date, end_date, trip_filter, carrier_filter=None):
    print_filter_summary(stats, start_date, end_date, trip_filter, carrier_filter)

    if not loads:
        print("Warning: No loads matched your filters - nothing to write.",
              file=sys.stderr)
        return

    # When you've filtered down to specific trips with --trips, each trip
    # gets its own tab in the workbook. Otherwise (no trip filter) sheets
    # are grouped by lane, same as before.
    group_by = 'trip' if trip_filter else 'lane'

    write_excel(loads, output_file, group_by=group_by)
    print(f"Successfully saved Excel report to '{output_file}'")
    print(
        "OT % columns are computed values (not live formulas), so they "
        "display correctly immediately in any viewer - no need to open "
        "and re-save through Excel/LibreOffice first."
    )


def resolve_date_filters(args):
    """
    Turns --this-week / --start-date / --end-date into a (start_date,
    end_date) pair of date objects (or None). --this-week sets the
    Mon-Sun range for today; explicit --start-date/--end-date override
    either side of it.
    """
    start_date = end_date = None

    if args.this_week:
        today = datetime.now().date()
        start_date = today - timedelta(days=today.weekday())  # Monday
        end_date = start_date + timedelta(days=6)              # Sunday

    if args.start_date:
        start_date = datetime.strptime(args.start_date, '%Y-%m-%d').date()

    if args.end_date:
        end_date = datetime.strptime(args.end_date, '%Y-%m-%d').date()

    return start_date, end_date


def resolve_trip_filter(args):
    return parse_id_list(args.trips)


def resolve_carrier_filter(args):
    if not args.carrier:
        return None
    return {c.strip() for c in args.carrier.split(',') if c.strip()}


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description='Build the trip on-time-performance Excel report.'
    )
    parser.add_argument(
        'input_file', nargs='?', default='input.txt',
        help='Local tab-delimited export to read (ignored when --sheet is used). Default: input.txt',
    )
    parser.add_argument(
        'output_file', nargs='?', default='output.xlsx',
        help='Path to write the .xlsx report to. Default: output.xlsx',
    )
    parser.add_argument(
        '--sheet', metavar='URL_OR_KEY',
        help='Read directly from a Google Sheet instead of a local file.',
    )
    parser.add_argument(
        '--worksheet', metavar='NAME',
        help='Worksheet/tab name when using --sheet (defaults to the first sheet).',
    )
    parser.add_argument(
        '--start-date', metavar='YYYY-MM-DD',
        help='Only include loads with a PU Date on/after this date.',
    )
    parser.add_argument(
        '--end-date', metavar='YYYY-MM-DD',
        help='Only include loads with a PU Date on/before this date.',
    )
    parser.add_argument(
        '--this-week', action='store_true',
        help='Shortcut: only include loads with a PU Date in the current Mon-Sun week.',
    )
    parser.add_argument(
        '--trips', metavar='LIST',
        help='Comma-separated Trip # and/or Load # to include, e.g. --trips 53,55,121744191',
    )
    parser.add_argument(
        '--carrier', metavar='LIST',
        help='Comma-separated carrier name(s) to include (matched case-insensitively), '
             'e.g. --carrier "Acme Trucking,Beta Logistics"',
    )
    return parser


def main():
    args = build_arg_parser().parse_args()

    output_file = args.output_file
    if not output_file.lower().endswith('.xlsx'):
        output_file = output_file.rsplit('.', 1)[0] + '.xlsx'

    start_date, end_date = resolve_date_filters(args)
    trip_filter = resolve_trip_filter(args)
    carrier_filter = resolve_carrier_filter(args)

    if args.sheet:
        loads, stats = load_trips_from_sheet(
            args.sheet, args.worksheet,
            start_date=start_date, end_date=end_date, trip_filter=trip_filter,
            carrier_filter=carrier_filter,
        )
    else:
        try:
            loads, stats = load_trips(
                args.input_file,
                start_date=start_date, end_date=end_date, trip_filter=trip_filter,
                carrier_filter=carrier_filter,
            )
        except FileNotFoundError:
            print(f"Error: Input file '{args.input_file}' not found!", file=sys.stderr)
            sys.exit(1)

    finish(loads, stats, output_file, start_date, end_date, trip_filter, carrier_filter)


if __name__ == '__main__':
    main()