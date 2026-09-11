"""
Streamlit web app for build_trip_report.py.

Run with:
    streamlit run app.py

Connects to your Google Sheet, lets you pick a PU date range / carrier /
trip selection and click Apply Filters, then shows the on-time % summary
on screen and offers the styled .xlsx report for download - all without
touching the command line.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st

from build_trip_report import (
    compute_on_time_pct,
    fetch_sheet_rows,
    parse_rows,
    write_excel,
)

st.set_page_config(page_title="Trip Report Builder", page_icon="🚚", layout="wide")


# ------------------------------------------------------------------
# Password gate. Only enforced when an app_password secret is actually
# configured (i.e. on the hosted deployment) - running locally with no
# .streamlit/secrets.toml just skips straight through, so this never
# locks you out of your own machine.
# ------------------------------------------------------------------

def _get_secret(key):
    try:
        return st.secrets.get(key)
    except Exception:
        return None


def _check_password(required_password):
    if st.session_state.get("password_ok"):
        return True

    st.title("🚚 Trip On-Time Report Builder")
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == required_password:
            st.session_state["password_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


_required_password = _get_secret("app_password")
if _required_password and not _check_password(_required_password):
    st.stop()


st.title("🚚 Trip On-Time Report Builder")
st.caption(
    "Connect your Google Sheet, filter by date / carrier / trip, and "
    "download the formatted Excel report."
)


# ------------------------------------------------------------------
# Step 1: connect to the sheet. There's only one sheet in use, so its
# URL lives in secrets (kept out of the public repo, same reasoning as
# supplier_name/contract_id) and the app connects to it automatically -
# no manual entry needed. Data is fetched once per session and cached,
# so every filter change below just re-slices the same data locally
# instead of re-hitting Google Sheets; use "Refresh data" to pull the
# latest data on demand. If sheet_url isn't configured (e.g. local use
# without secrets set up), a manual entry box is shown instead.
# ------------------------------------------------------------------

_default_sheet_url = _get_secret("sheet_url")
_default_worksheet = _get_secret("worksheet_name")  # optional, rarely needed

def _invalidate_data():
    st.session_state.pop("raw_rows", None)
    st.session_state.pop("all_loads", None)
    st.session_state.pop("all_stats", None)


with st.sidebar:
    if _default_sheet_url:
        if st.button("🔄 Refresh data", use_container_width=True):
            _invalidate_data()
    else:
        st.subheader("🔗 Connect")
        st.caption("No sheet_url configured in secrets - enter one manually.")
        sheet_input = st.text_input(
            "Google Sheet URL or key",
            value=st.session_state.get("sheet_input", ""),
        )
        worksheet_input = st.text_input(
            "Worksheet/tab name (optional)",
            value=st.session_state.get("worksheet_input", ""),
        )
        if st.button("Connect / Refresh data", type="primary", use_container_width=True):
            st.session_state["sheet_input"] = sheet_input.strip()
            st.session_state["worksheet_input"] = worksheet_input.strip()
            _invalidate_data()

_active_sheet_url = _default_sheet_url or st.session_state.get("sheet_input")
_active_worksheet = _default_worksheet or st.session_state.get("worksheet_input") or None

if "raw_rows" not in st.session_state:
    if not _active_sheet_url:
        st.info("👈 Enter your Google Sheet to get started.")
        st.stop()
    with st.spinner("Reading sheet..."):
        try:
            st.session_state["raw_rows"] = fetch_sheet_rows(_active_sheet_url, _active_worksheet)
        except Exception as e:
            st.error(f"Couldn't read the sheet:\n\n{e}")
            st.stop()

raw_rows = st.session_state["raw_rows"]
st.sidebar.caption(f"✅ Connected · {max(len(raw_rows) - 1, 0)} rows loaded")

# Parse the sheet into loads ONCE per connection and cache it. This is
# the expensive step (cleaning + datetime-parsing every field, for every
# row) - doing it here instead of on every filter tweak is what actually
# makes filtering fast, since adjusting filters below then only ever
# re-slices this already-parsed list rather than re-reading the sheet.
if "all_loads" not in st.session_state:
    try:
        st.session_state["all_loads"], st.session_state["all_stats"] = parse_rows(raw_rows)
    except ValueError as e:
        st.error(str(e))
        st.stop()

all_loads = st.session_state["all_loads"]
_all_stats = st.session_state["all_stats"]

if not all_loads:
    st.warning("That sheet has no usable rows (after excluding blank/cancelled/TONU loads).")
    st.stop()


# ------------------------------------------------------------------
# Step 2: filters
# ------------------------------------------------------------------

pu_dates = [ld["sched_arr"].date() for ld in all_loads if ld["sched_arr"]]
data_min_date = min(pu_dates) if pu_dates else datetime.now().date()
data_max_date = max(pu_dates) if pu_dates else datetime.now().date()

all_carriers = sorted({ld["carrier"] for ld in all_loads})

# Numeric sort (1, 2, 6, 8, 34, ...) instead of string sort (1, 108, 2, ...)
# so the picker below lists trips in the order you'd actually expect.
all_trip_ids = sorted(
    {ld["trip_id"] for ld in all_loads if ld["trip_id"]},
    key=lambda t: (0, int(t)) if t.isdigit() else (1, t),
)

with st.sidebar:
    st.divider()
    st.subheader("🎚️ Filters")

    # Clamp any remembered date_range into the current data's bounds
    # before handing it to the widget below - a stale value from a
    # previous sheet/connection (or the actual data range shifting once
    # filters like "regular trips only" are applied) would otherwise
    # crash date_input with a value outside its own min/max.
    _default_range = st.session_state.get("date_range", (data_min_date, data_max_date))
    if isinstance(_default_range, tuple) and len(_default_range) == 2:
        _clamped_start = min(max(_default_range[0], data_min_date), data_max_date)
        _clamped_end = min(max(_default_range[1], data_min_date), data_max_date)
        st.session_state["date_range"] = (_clamped_start, _clamped_end)
    else:
        st.session_state["date_range"] = (data_min_date, data_max_date)

    # Everything in this form is staged, not live - picking dates,
    # (de)selecting carriers/trips doesn't touch the report below at all
    # until "Apply Filters" is clicked, which is the one moment
    # everything recomputes.
    with st.form("filters_form"):
        date_range = st.date_input(
            "PU Date range",
            min_value=data_min_date,
            max_value=data_max_date,
            key="date_range",
        )

        selected_carriers = st.multiselect(
            "Carrier (default = all)",
            options=all_carriers,
            default=all_carriers,
        )

        selected_trips = st.multiselect(
            "Trip # (default = all regular trips)",
            options=all_trip_ids,
            default=all_trip_ids,
            help="Deselect the trips you don't want in the report, same as Carrier above. Type to search.",
        )

        st.form_submit_button("🔍 Apply Filters", type="primary", use_container_width=True)

# Resolve widget values into parse_rows()-compatible filters.
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    # date_input can briefly return a single date while the user is
    # picking the second end of the range.
    start_date = end_date = date_range

# Only treat a selection as a real filter once it's a strict subset - if
# everything is (still) selected, that's equivalent to no filter (and
# for trips specifically, "no filter" already means "regular trips
# only" by default - see REGULAR_TRIP_IDS in build_trip_report.py).
carrier_filter = set(selected_carriers) if set(selected_carriers) != set(all_carriers) else None
trip_filter = set(selected_trips) if set(selected_trips) != set(all_trip_ids) else None


def filter_loads(loads, start_date, end_date, carrier_filter, trip_filter):
    """
    Re-filters an already-parsed loads list (mirrors the equivalent
    checks in build_trip_report.parse_rows(), just against pre-parsed
    dicts instead of raw sheet strings) - cheap date/set comparisons
    only, no re-cleaning or re-parsing, so this is fast enough to run on
    every "Apply Filters" click.
    """
    stats = {
        'skipped_unparseable_date': 0,
        'skipped_date_range': 0,
        'skipped_carrier_filter': 0,
        'skipped_trip_filter': 0,
        'included': 0,
        'total_rows': len(loads),
    }
    result = []
    for ld in loads:
        if start_date is not None or end_date is not None:
            if ld['sched_arr'] is None:
                stats['skipped_unparseable_date'] += 1
                continue
            day = ld['sched_arr'].date()
            if start_date is not None and day < start_date:
                stats['skipped_date_range'] += 1
                continue
            if end_date is not None and day > end_date:
                stats['skipped_date_range'] += 1
                continue
        if carrier_filter is not None and ld['carrier'] not in carrier_filter:
            stats['skipped_carrier_filter'] += 1
            continue
        if trip_filter is not None and ld['trip_id'] not in trip_filter:
            stats['skipped_trip_filter'] += 1
            continue
        result.append(ld)
        stats['included'] += 1
    return result, stats


loads, stats = filter_loads(all_loads, start_date, end_date, carrier_filter, trip_filter)


# ------------------------------------------------------------------
# Step 3: on-screen summary
# ------------------------------------------------------------------

st.divider()
st.header("📊 Summary")

# Filters apply live (every change above reruns this instantly - there's
# no separate "Apply" step), but that's not obvious just from the
# sidebar, so spell out exactly what's active and how many loads it
# produced right here, every time.
_carrier_note = "all carriers" if carrier_filter is None else f"{len(carrier_filter)} carrier(s) selected"
_trip_note = "regular trips only" if trip_filter is None else f"{len(trip_filter)} specific trip(s)"
st.success(
    f"✅ **Filters applied** — {start_date} → {end_date} · {_carrier_note} · {_trip_note} "
    f"→ **{len(loads)}** loads match"
)

if not loads:
    st.warning("No loads match the current filters.")
    st.stop()

summary_box = st.container(border=True)
col1, col2, col3, col4 = summary_box.columns(4)
col1.metric("Loads", len(loads))


def pct_label(loads, field):
    pct = compute_on_time_pct(loads, field)
    return "n/a" if pct is None else f"{pct:.0%}"


def pct100(loads, field):
    """
    compute_on_time_pct() returns a 0-1 fraction (Excel's '0%' cell format
    multiplies by 100 automatically when displaying it, which is why the
    .xlsx report always looked right). Streamlit's NumberColumn format
    string does NOT auto-multiply - "%.0f%%" on a raw fraction like 0.885
    just rounds *that* to "1" and appends a "%", which is why the on-screen
    tables were showing "1%" almost everywhere. Scaling to 0-100 here
    before it reaches st.dataframe() is what actually fixes it.
    """
    pct = compute_on_time_pct(loads, field)
    return None if pct is None else pct * 100


col2.metric("OT Arrival", pct_label(loads, "on_time_arr"))
col3.metric("OT Dispatch", pct_label(loads, "on_time_dispatch"))
col4.metric("OT Delivery", pct_label(loads, "on_time_delivery"))

with st.expander(f"Row filtering detail ({stats['included']} of {stats['total_rows']} regular-trip rows shown)"):
    st.write(
        "**Always excluded (data quality, from the full sheet):**\n"
        f"- Blank rows: {_all_stats['skipped_blank']}\n"
        f"- Cancelled/TONU: {_all_stats['skipped_cancelled_tonu']}\n"
        f"- No check-in/out at pickup or delivery: {_all_stats['skipped_missing_checkin']}\n"
        f"- Not a regular trip: {_all_stats['skipped_trip_filter']}\n"
        "\n**Excluded by your current filters:**\n"
        f"- Unparseable PU date: {stats['skipped_unparseable_date']}\n"
        f"- Outside date range: {stats['skipped_date_range']}\n"
        f"- Carrier filter: {stats['skipped_carrier_filter']}\n"
        f"- Trip filter: {stats['skipped_trip_filter']}"
    )

tab_carrier, tab_lane, tab_preview = st.tabs(["By Carrier", "By Lane", "Load preview"])

with tab_carrier:
    by_carrier = {}
    for ld in loads:
        by_carrier.setdefault(ld["carrier"], []).append(ld)
    carrier_df = pd.DataFrame(
        [
            {
                "Carrier": name,
                "Loads": len(group),
                "OT Arrival %": pct100(group, "on_time_arr"),
                "OT Dispatch %": pct100(group, "on_time_dispatch"),
                "OT Delivery %": pct100(group, "on_time_delivery"),
            }
            for name, group in sorted(by_carrier.items())
        ]
    )
    st.dataframe(
        carrier_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            c: st.column_config.NumberColumn(format="%.0f%%")
            for c in ("OT Arrival %", "OT Dispatch %", "OT Delivery %")
        },
    )

with tab_lane:
    by_lane = {}
    for ld in loads:
        by_lane.setdefault(ld["lane"], []).append(ld)
    lane_df = pd.DataFrame(
        [
            {
                "Lane": name,
                "Loads": len(group),
                "OT Arrival %": pct100(group, "on_time_arr"),
                "OT Dispatch %": pct100(group, "on_time_dispatch"),
                "OT Delivery %": pct100(group, "on_time_delivery"),
            }
            for name, group in sorted(by_lane.items())
        ]
    )
    st.dataframe(
        lane_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            c: st.column_config.NumberColumn(format="%.0f%%")
            for c in ("OT Arrival %", "OT Dispatch %", "OT Delivery %")
        },
    )

with tab_preview:
    preview_df = pd.DataFrame(
        [
            {
                "Trip ID": ld["trip_id"],
                "Load ID": ld["load_id"],
                "Carrier": ld["carrier"],
                "Lane": ld["lane"],
                "Sched Arrival": ld["sched_arr"],
                "Actual Arrival": ld["actual_arr"],
                "OT Arrival": ld["on_time_arr"],
                "OT Dispatch": ld["on_time_dispatch"],
                "OT Delivery": ld["on_time_delivery"],
                "Notes": ld["notes"],
            }
            for ld in loads
        ]
    )
    st.dataframe(preview_df, use_container_width=True, hide_index=True)


# ------------------------------------------------------------------
# Step 4: download
# ------------------------------------------------------------------

st.divider()
st.header("📥 Download")

group_by = "trip" if trip_filter else "lane"
buffer = io.BytesIO()
write_excel(loads, buffer, group_by=group_by)
buffer.seek(0)

date_part = f"{start_date or 'earliest'}_to_{end_date or 'latest'}"
file_name = f"trip_report_{date_part}.xlsx"

st.download_button(
    "⬇️ Download Excel report",
    data=buffer,
    file_name=file_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.divider()
st.caption("Made by Lamarr Glover")
