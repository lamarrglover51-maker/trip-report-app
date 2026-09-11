"""
Streamlit web app for build_trip_report.py.

Run with:
    streamlit run app.py

Lets you connect to your Google Sheet once, then interactively filter by
PU date range, carrier, and trip/load number, see the on-time % summary
right on screen, and download the styled .xlsx report - all without
touching the command line.
"""

import io
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

from build_trip_report import (
    compute_on_time_pct,
    fetch_sheet_rows,
    parse_id_list,
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
    "download the formatted Excel report. · LG"
)


# ------------------------------------------------------------------
# Step 1: connect to the sheet (fetched once, cached in session_state so
# every filter change below just re-slices the same data locally instead
# of re-hitting Google Sheets).
# ------------------------------------------------------------------

with st.sidebar:
    st.subheader("🔗 Connect")
    sheet_input = st.text_input(
        "Google Sheet URL or key",
        value=st.session_state.get("sheet_input", ""),
        help="Paste the full sheet URL, or just the sheet's key/ID.",
    )
    worksheet_input = st.text_input(
        "Worksheet/tab name (optional)",
        value=st.session_state.get("worksheet_input", ""),
        help="Leave blank to use the first tab in the sheet.",
    )
    connect_clicked = st.button("Connect / Refresh data", type="primary", use_container_width=True)

if connect_clicked:
    if not sheet_input.strip():
        st.sidebar.error("Enter a Google Sheet URL or key first.")
    else:
        with st.spinner("Reading sheet... (a browser window may open the first time, for Google login)"):
            try:
                rows = fetch_sheet_rows(sheet_input.strip(), worksheet_input.strip() or None)
                st.session_state["raw_rows"] = rows
                st.session_state["sheet_input"] = sheet_input.strip()
                st.session_state["worksheet_input"] = worksheet_input.strip()
                st.sidebar.success(f"Loaded {max(len(rows) - 1, 0)} rows.")
            except Exception as e:
                st.sidebar.error(f"Couldn't read that sheet:\n\n{e}")

if "raw_rows" not in st.session_state:
    st.info("👈 Enter your Google Sheet and click **Connect / Refresh data** to get started.")
    st.stop()

raw_rows = st.session_state["raw_rows"]

# Parse once with no filters, purely to populate the filter widgets
# (unique carriers, date bounds) and to know the full unfiltered size.
try:
    all_loads, _all_stats = parse_rows(raw_rows)
except ValueError as e:
    st.error(str(e))
    st.stop()

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

with st.sidebar:
    st.divider()
    st.subheader("🎚️ Filters")

    if st.button("📅 This week (Mon-Sun)", use_container_width=True):
        today = datetime.now().date()
        st.session_state["date_range"] = (
            today - timedelta(days=today.weekday()),
            today - timedelta(days=today.weekday()) + timedelta(days=6),
        )

    date_range = st.date_input(
        "PU Date range",
        value=st.session_state.get("date_range", (data_min_date, data_max_date)),
        min_value=data_min_date,
        max_value=data_max_date,
        key="date_range",
    )

    selected_carriers = st.multiselect(
        "Carrier (default = all)",
        options=all_carriers,
        default=all_carriers,
    )

    trips_text = st.text_area(
        "Trip # / Load # allowlist (optional)",
        height=100,
        help=(
            'Only these trips will be included - anything not in this list is '
            'ignored. Comma-separated ("53,55,121744191") or one per line '
            '(paste a column straight out of Excel). Leave blank to include '
            'every trip.'
        ),
    )

# Resolve widget values into parse_rows()-compatible filters.
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    # date_input can briefly return a single date while the user is
    # picking the second end of the range.
    start_date = end_date = date_range

trip_filter = parse_id_list(trips_text)
# Only treat carrier selection as a real filter once it's a strict subset
# - if everything is (still) selected, that's equivalent to no filter.
carrier_filter = set(selected_carriers) if set(selected_carriers) != set(all_carriers) else None

loads, stats = parse_rows(
    raw_rows,
    start_date=start_date,
    end_date=end_date,
    trip_filter=trip_filter,
    carrier_filter=carrier_filter,
)


# ------------------------------------------------------------------
# Step 3: on-screen summary
# ------------------------------------------------------------------

st.divider()
st.header("📊 Summary")

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

with st.expander(f"Row filtering detail ({stats['included']} of {stats['total_rows']} rows included)"):
    st.write(
        f"- Skipped (blank row): {stats['skipped_blank']}\n"
        f"- Skipped (cancelled/TONU): {stats['skipped_cancelled_tonu']}\n"
        f"- Skipped (outside date range): {stats['skipped_date_range']}\n"
        f"- Skipped (unparseable date): {stats['skipped_unparseable_date']}\n"
        f"- Skipped (trip filter): {stats['skipped_trip_filter']}\n"
        f"- Skipped (carrier filter): {stats['skipped_carrier_filter']}\n"
        f"- Skipped (no check-in/out at pickup or delivery): {stats['skipped_missing_checkin']}"
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
