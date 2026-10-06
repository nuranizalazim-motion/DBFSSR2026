"""Faculty dashboard. All analytical displays derive from one filtered frame."""
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo
import json
import os

import pandas as pd
import plotly.express as px
import streamlit as st

from data_pipeline import apply_filters, build_snapshot, read_sources, refresh_snapshot, safe_export
from drive_source import read_drive_sources

ROOT = Path(__file__).resolve().parent
DATA = Path(os.environ.get("DASHBOARD_DATA_DIR", ROOT / "data"))
COLOURS = ["#4054A8", "#187E79", "#AA6816", "#8A439B", "#BB4158", "#596575"]
DIMENSIONS = {
    "Students": ["Campus", "Programme"],
    "Programmes": ["Campus", "Code group", "Review status", "Review stage"],
    "Staff planning": ["Department", "Grade", "Qualification label", "Possible duplicate name"],
    "Directory": ["Staff category", "Section"],
    "WBL groups": ["Programme", "Semester (source)", "Group"],
    "Historical intake": ["Programme", "Campus", "Year", "Session"],
    "Facilities": ["Campus", "Block", "Floor (source)", "Source file"],
}
NOTES = {
    "Students": "Degree 2026 snapshot · Counts are students, not individual student records. Blank campus cells remain missing. Campus definitions come from the historical sheet. This is separate from the WBL and programme-report snapshots.",
    "Programmes": "Programme + campus is the verified join key. Each code group lists aliases for one programme; aliases do not create extra records. Missing counts are not zero. Ratios use the programme report only and require both staff count fields.",
    "Staff planning": "Counts represent source records, not confirmed unique people. Repeated names are flagged, retained and never joined by fuzzy name matching. Qualifications are what the source mentions; awards/completion are not verified. Ages are from source snapshots, not recalculated.",
    "Directory": "Directory entries include management and departmental listings. One person can have multiple entries. These are not added to the staff planning count or joined to qualifications.",
    "WBL groups": "Session 20262 · Mac–Ogos 2026. This snapshot differs from Degree 2026. Explicit addition 17 + 1 becomes 18; 2 (CK) stays unparsed and excluded from the numeric sum. Repeated group IDs are retained for review.",
    "Historical intake": "Intake registrations are not current enrolment. Choose programme totals or campus breakdowns; these two levels are never combined. Some campus sums do not reconcile to their parent totals. UNJ = projected; DFT = registered; TWR = offered. Blank and dash cells remain missing.",
    "Facilities": "PDF: individual Puncak Alam rooms. PowerPoint: Shah Alam room groups. Capacity is the reported total for each row/group, never multiplied by the room count. Unknown student capacities remain missing; staff and general-person capacities are kept separate. No timetable or occupancy data was supplied.",
}

st.set_page_config(page_title="FSSR Faculty Insights", page_icon="🎨", layout="wide")
st.markdown("""<style>
.stApp {background:#F7F8FC;} [data-testid='stSidebar'] {background:#FFFFFF;}
h1,h2,h3 {color:#202440;} [data-testid='stMetric'] {background:white;border:1px solid #E1E4EF;border-radius:14px;padding:18px;}
.eyebrow {font-size:12px;letter-spacing:2px;color:#635A89;font-weight:700;}
.hero {background:#24253F;color:white;border-radius:18px;padding:28px 32px;margin-bottom:22px;}
.hero h1 {color:white;margin:4px 0 8px;font-size:36px;} .hero p {color:#E0DEEF;margin:0;}
</style><div class='hero'><div class='eyebrow' style='color:#E7C787'>UiTM · FACULTY OF ART AND DESIGN</div>
<h1>FSSR Faculty Insights</h1><p>Students, people, programmes and spaces — a shared view for faculty decisions.</p></div>""", unsafe_allow_html=True)


try:
    DRIVE_CONFIG = dict(st.secrets.get("google_drive", {}))
    SECRETS_ERROR = None
except FileNotFoundError:
    DRIVE_CONFIG = {}
    SECRETS_ERROR = None
except Exception:
    DRIVE_CONFIG = {}
    SECRETS_ERROR = "Cannot read Streamlit Secrets. Check the TOML formatting in the app's Secrets settings."


def load_current(replacements=None):
    if SECRETS_ERROR:
        raise ValueError(SECRETS_ERROR)
    if DRIVE_CONFIG:
        sources, drive_manifest = read_drive_sources(DRIVE_CONFIG)
        candidate = build_snapshot(sources)
        by_name = {item["File"]: item for item in drive_manifest}
        for item in candidate["manifest"]:
            item.update(by_name[item["File"]])
        candidate["source_mode"] = "Google Drive"
        return candidate
    candidate = build_snapshot(read_sources(DATA, replacements))
    candidate["source_mode"] = "Session uploads" if replacements else "Local files"
    return candidate


if "snapshot" not in st.session_state and "initial_read_attempted" not in st.session_state and (DRIVE_CONFIG or SECRETS_ERROR or DATA.is_dir()):
    st.session_state["initial_read_attempted"] = True
    with st.spinner("Reading and validating source files…"):
        refresh_snapshot(st.session_state, load_current)

with st.sidebar:
    st.markdown("### Explore your faculty")
    view = st.radio("Data view", list(DIMENSIONS), key="view")
    st.divider()
    st.markdown("### Source refresh")
    st.caption("Filters update immediately. Refresh rereads source files and runs validation.")
    if st.button("Refresh source files", width="stretch"):
        with st.spinner("Reading and validating source files…"):
            refresh_snapshot(st.session_state, lambda: load_current(st.session_state.get("replacements")))
    with st.expander("Replace source files"):
        if DRIVE_CONFIG:
            st.caption("Shared source files come from your private Google Drive folder. Update the files there, keep their original filenames, then press Refresh source files. The app has read-only access.")
        else:
            st.caption("Upload revised files with their original filenames. Uploaded replacements apply to this session only. Originals stay unchanged.")
            uploads = st.file_uploader("Updated Excel / PDF / PowerPoint files", type=["xlsx", "pdf", "pptx"], accept_multiple_files=True)
            if st.button("Validate and apply uploads", disabled=not uploads):
                replacements = dict(st.session_state.get("replacements", {}))
                for u in uploads:
                    replacements[u.name] = u.getvalue()
                if refresh_snapshot(st.session_state, lambda: load_current(replacements)):
                    st.session_state["replacements"] = replacements
            if st.button("Use local source folder"):
                if refresh_snapshot(st.session_state, load_current):
                    st.session_state.pop("replacements", None)
    stale_hours = st.number_input("Flag data after (hours since read)", min_value=1, value=24)
    if DRIVE_CONFIG:
        st.caption("Source: private Google Drive. Loads when a new session opens; use Refresh source files for updates. No background refresh schedule is active.")
    else:
        st.caption("Manual refresh is active. Google Drive is not connected. Permanent setup: see GOOGLE_DRIVE_SETUP.md in the GitHub repository.")

if st.session_state.get("refresh_error"):
    st.error("Refresh failed: " + st.session_state["refresh_error"] + ". Showing the last successful data if available.")
snapshot = st.session_state.get("snapshot")
if snapshot is None:
    if DRIVE_CONFIG:
        st.info("Google Drive is configured, but no source read has succeeded yet. Follow the error above, then press Refresh source files to retry.")
    else:
        st.info("This installation has no faculty data yet. Open Replace source files in the sidebar, upload all 11 original files using their original filenames, then press Validate and apply uploads. These uploads are temporary; connect private Google Drive to load sources again after a session restarts.")
    st.caption("For a shared website, configure private access in your hosting settings before uploading faculty records.")
    st.stop()
loaded = datetime.fromisoformat(snapshot["loaded_at"])
local = loaded.astimezone(ZoneInfo("Asia/Kuala_Lumpur"))
st.caption(f"Last successful source read: {local:%d %b %Y, %I:%M:%S %p} MYT · {len(snapshot['manifest'])} source files · Source: {snapshot.get('source_mode', 'Local files')} · Filter changes do not refresh the files")
if (datetime.now(timezone.utc) - loaded).total_seconds() > stale_hours * 3600:
    st.warning("Stale data: the last successful source read is older than your chosen threshold.")
st.caption("The original files do not establish one shared update date. A successful read does not prove the underlying records are current.")
st.subheader(view)
st.info(NOTES[view])
frame = snapshot["tables"][view]
prefix = f"filter:{view}:"
if view == "Historical intake":
    level = st.selectbox("Historical record level", ["Programme totals", "Campus breakdown"], key=prefix + "level")
    frame = frame[frame["Record level"] == level]


def reset_filters():
    for key in list(st.session_state):
        if key.startswith(prefix):
            del st.session_state[key]


with st.sidebar:
    st.divider()
    st.markdown("### Filter this view")
    st.button("Reset filters", on_click=reset_filters, width="stretch")
    categories = {}
    for column in DIMENSIONS[view]:
        values = sorted(frame[column].fillna("Not supplied").astype(str).unique())
        key = prefix + column
        # Retain active selections across refresh, including removed categories.
        old = st.session_state.get(key, values)
        options = sorted(set(values) | set(old))
        categories[column] = st.multiselect(column, options, default=old, key=key)
    query = st.text_input("Search all record fields", key=prefix + "search")
    date_filter = number_filter = None
    if view == "Programmes":
        date_field = st.selectbox("Date field", ["Maturity date", "Programme start", "Accreditation start", "Accreditation end", "Review date"], key=prefix + "datefield")
        use_dates = st.checkbox("Apply a date range", key=prefix + "usedate")
        dates = frame[date_field].dropna()
        if use_dates and not dates.empty:
            date_range = st.date_input("Date range (inclusive)", value=(dates.min().date(), dates.max().date()), key=prefix + "dates:" + date_field)
            if len(date_range) == 2:
                date_filter = (date_field, *date_range)
                st.caption("Records without this date are excluded while the range is active.")
            else:
                st.caption("Choose an end date to apply the range.")
        elif use_dates:
            st.caption("No dates supplied for this field. No date restriction applied.")
    elif view == "Staff planning" and st.checkbox("Filter retirement years", key=prefix + "useyears"):
        years = frame["Retirement year"].dropna()
        if len(years) and years.min() < years.max():
            bounds = st.slider("Retirement year range", int(years.min()), int(years.max()), (int(years.min()), int(years.max())), key=prefix + "years")
            number_filter = ("Retirement year", *bounds)
            st.caption("Missing retirement years are excluded while this filter is active.")

filtered = apply_filters(frame, categories, query, date_filter, number_filter)
st.caption(f"{len(filtered):,} of {len(frame):,} records selected · All figures, charts, tables and exports below use this selection.")


def metrics(items):
    for col, (label, value) in zip(st.columns(len(items)), items):
        col.metric(label, value)


def total(series):
    n = series.sum(min_count=1)
    return "Not supplied" if pd.isna(n) else f"{n:,.0f}"


charts = []


def show_chart(figure, name):
    figure.update_layout(font=dict(family="Arial", size=13, color="#202440"), paper_bgcolor="white", plot_bgcolor="white",
                         margin=dict(l=18, r=18, t=60, b=30), legend_title_text="", title_font_size=18)
    st.plotly_chart(figure, use_container_width=True, config={"displaylogo": False, "toImageButtonOptions": {"format": "png", "filename": name}})
    charts.append((name, figure))


def counts_chart(column, title):
    grouped = filtered[column].fillna("Not supplied").value_counts().rename_axis(column).reset_index(name="Records")
    fig = px.bar(grouped, x="Records", y=column, orientation="h", title=title, color_discrete_sequence=COLOURS, text="Records")
    fig.update_layout(height=max(350, min(1100, len(grouped) * 32 + 100)))
    show_chart(fig, title)


if filtered.empty:
    st.warning("No records match. Change your selections or use Reset filters.")
elif view == "Students":
    metrics([("Students in nonblank cells", total(filtered.Students)), ("Programmes selected", filtered.Programme.nunique()),
             ("Campuses selected", filtered.Campus.nunique()), ("Missing campus counts", int(filtered.Students.isna().sum()))])
    programme = filtered.groupby("Programme", as_index=False).Students.sum(min_count=1).dropna()
    campus = filtered.groupby("Campus", as_index=False).Students.sum(min_count=1).dropna()
    show_chart(px.bar(programme, x="Students", y="Programme", orientation="h", title="Students by programme", color_discrete_sequence=COLOURS), "students-programme")
    show_chart(px.bar(campus, x="Campus", y="Students", title="Students by campus", text="Students", color_discrete_sequence=COLOURS), "students-campus")
elif view == "Programmes":
    metrics([("Programme-campus records", len(filtered)), ("Review marked Selesai", int((filtered["Review status"] == "Selesai").sum())),
             ("Student counts supplied", int(filtered["Students (report source)"].notna().sum())),
             ("Maturity dates supplied", int(filtered["Maturity date"].notna().sum()))])
    counts_chart("Review status", "Programme review status")
    ratios = filtered.dropna(subset=["Students per reported staff"])
    if not ratios.empty:
        show_chart(px.bar(ratios, x="Students per reported staff", y="Code group", color="Campus", orientation="h",
                          title="Student-to-staff ratio · programme report only", hover_data=["Students (report source)", "Staff total (complete pairs only)"],
                          color_discrete_sequence=COLOURS), "programme-ratios")
    st.caption("Maturity dates are shown as dates, not treated as accreditation expiry. No expiry classification is inferred from missing end dates.")
elif view == "Staff planning":
    metrics([("Staff source records", len(filtered)), ("PhD mentioned in source", int((filtered["Qualification label"] == "PhD mentioned").sum())),
             ("Repeated-name records", int(filtered["Possible duplicate name"].sum())), ("Missing retirement years", int(filtered["Retirement year"].isna().sum()))])
    counts_chart("Department", "Staff records by department")
    counts_chart("Qualification label", "Qualifications mentioned in the source")
    years = filtered["Retirement year"].dropna().astype(int).value_counts().sort_index().rename_axis("Year").reset_index(name="Records")
    if not years.empty:
        show_chart(px.bar(years, x="Year", y="Records", title="Reported retirement years", color_discrete_sequence=COLOURS, text="Records"), "retirement-years")
elif view == "Directory":
    metrics([("Directory entries", len(filtered)), ("Sections selected", filtered.Section.nunique()),
             ("Repeated-name entries", int(filtered["Possible duplicate name"].sum())), ("Missing phone numbers", int(filtered["Phone (source)"].isna().sum()))])
    counts_chart("Staff category", "Directory entries by staff category")
    counts_chart("Section", "Directory entries by section")
elif view == "WBL groups":
    metrics([("Unambiguous group count sum", total(filtered.Students)), ("Group records", len(filtered)),
             ("Unparsed count cells", int(filtered.Students.isna().sum())), ("Repeated group rows", int(filtered.Group.duplicated(keep=False).sum()))])
    show_chart(px.bar(filtered, x="Group", y="Students", color="Programme", title="WBL student counts by group", hover_data=["Count (source)", "Semester (source)"], color_discrete_sequence=COLOURS), "wbl-groups")
elif view == "Historical intake":
    metrics([("Registrations across selected intakes", total(filtered["Registered (DFT)"])), ("Projected across selected intakes", total(filtered["Projected (UNJ)"])),
             ("Sessions selected", filtered.Session.nunique()), ("Missing registration cells", int(filtered["Registered (DFT)"].isna().sum()))])
    periods = filtered.groupby("Period", as_index=False)[["Registered (DFT)", "Projected (UNJ)"]].sum(min_count=1)
    show_chart(px.line(periods, x="Period", y=["Registered (DFT)", "Projected (UNJ)"], markers=True,
                       title="Intake registrations and projections", color_discrete_sequence=COLOURS), "intake-trends")
elif view == "Facilities":
    metrics([("Reported rooms", total(filtered.Rooms)), ("Student capacity supplied", total(filtered["Student capacity (source)"])),
             ("Rows with unknown student capacity", int(filtered["Student capacity (source)"].isna().sum())), ("Source rows / groups", len(filtered))])
    blocks = filtered.groupby(["Campus", "Block"], as_index=False).Rooms.sum(min_count=1)
    show_chart(px.bar(blocks, x="Block", y="Rooms", color="Campus", barmode="group", title="Reported rooms by campus and block", color_discrete_sequence=COLOURS), "facility-rooms")
    capacity = filtered.groupby("Campus", as_index=False)["Student capacity (source)"].sum(min_count=1).dropna()
    show_chart(px.bar(capacity, x="Campus", y="Student capacity (source)", title="Reported student capacity · not occupancy", color_discrete_sequence=COLOURS), "facility-capacity")

st.markdown("### Records behind this view")
sort_column = st.selectbox("Sort records by", list(filtered.columns), key=prefix + "sort")
descending = st.checkbox("Descending order", key=prefix + "descending")
sorted_records = filtered.sort_values(sort_column, ascending=not descending, na_position="last", kind="stable")
st.dataframe(sorted_records, hide_index=True, width="stretch")
st.caption("Click a table column header for additional sorting. Use the search field in the sidebar to find records.")
export = safe_export(sorted_records)
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    export.to_excel(writer, index=False, sheet_name="Selected records")
c1, c2 = st.columns(2)
c1.download_button("Download selected data · CSV", export.to_csv(index=False).encode("utf-8-sig"), file_name=f"fssr-{view.lower().replace(' ', '-')}.csv", mime="text/csv")
c2.download_button("Download selected data · Excel", output.getvalue(), file_name=f"fssr-{view.lower().replace(' ', '-')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with st.expander("Export charts"):
    st.caption("Use each chart’s camera icon for PNG. HTML exports retain tooltips and zoom and work offline.")
    for chart_name, chart in charts:
        st.download_button(f"Download {chart_name} · HTML", chart.to_html(include_plotlyjs=True), file_name=f"{chart_name}.html", mime="text/html", key="chart:" + chart_name)
with st.expander("Data quality and source validation"):
    st.caption("These checks cover the complete imported sources, independently of the current view filters.")
    st.dataframe(snapshot["audit"], hide_index=True, width="stretch")
    st.download_button("Download validation report", snapshot["audit"].to_csv(index=False).encode(), "validation.csv", "text/csv")
    st.download_button("Download worksheet profile", json.dumps(snapshot["profiles"], indent=2), "source_profile.json", "application/json")
    st.dataframe(pd.DataFrame(snapshot["manifest"]), hide_index=True, width="stretch")
with st.expander("Help · updating and sharing"):
    st.markdown("""**Updating:** replace files in the local `data` folder using the same filenames, then press **Refresh source files**. Or apply revised files in **Replace source files** for this session. Filter selections stay active; use **Reset filters** if you want to start over.

**Google Drive:** with the private connection configured, new sessions read the shared Drive folder automatically. **Refresh source files** rereads it and validates all eleven files while preserving active filters and the current session's last successful data on failure. There is no background refresh schedule. Setup instructions are in **GOOGLE_DRIVE_SETUP.md** in the GitHub repository. Without this connection, browser uploads are temporary and local data folders are read manually.

**Sharing:** the local prototype has no login or management/lecturer access roles. Use an institution-managed host with authentication before granting shared access. Full setup and maintenance instructions are in `docs/BEGINNER_GUIDE.md`.""")
