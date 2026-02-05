import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from urllib.parse import quote
from zoneinfo import ZoneInfo
import uuid

st.set_page_config(page_title="Prospecting Manager", page_icon="📇", layout="wide")

# ----------------------------- Utility & Constants -----------------------------
REQUIRED_COLUMNS = ["Name", "Phone", "Email"]
OPTIONAL_COLUMNS = ["Company", "Title", "State"]
APP_COLUMNS = REQUIRED_COLUMNS + OPTIONAL_COLUMNS + [
    "Status", "MeetingDateTime", "CallbackDateTime", "LastCallDateTime", "Attempts", "Notes"
]

STATUS_OPTIONS = ["", "Voicemail", "Yes", "No", "Call back later"]

COMMON_TIMEZONES = [
    "America/New_York", "America/Chicago", "America/Denver", "America/Los_Angeles",
    "America/Phoenix", "America/Anchorage", "America/Honolulu", "UTC"
]

def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Normalize column names
    rename_map = {}
    cols_lower = {c.lower(): c for c in df.columns}
    for col in ["name", "phone", "email", "company", "title", "state", "status",
                "meetingdatetime", "callbackdatetime", "lastcalldatetime", "attempts", "notes"]:
        if col in cols_lower:
            src = cols_lower[col]
            if col in ["meetingdatetime", "callbackdatetime", "lastcalldatetime"]:
                rename_map[src] = col.capitalize().replace("datetime", "DateTime")
            elif col == "email":
                rename_map[src] = "Email"
            else:
                rename_map[src] = col.capitalize()
    df = df.rename(columns=rename_map)

    # Ensure required/optional/app columns exist
    for col in REQUIRED_COLUMNS + OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    if "Status" not in df.columns:
        df["Status"] = ""
    if "MeetingDateTime" not in df.columns:
        df["MeetingDateTime"] = pd.NaT
    if "CallbackDateTime" not in df.columns:
        df["CallbackDateTime"] = pd.NaT
    if "LastCallDateTime" not in df.columns:
        df["LastCallDateTime"] = pd.NaT
    if "Attempts" not in df.columns:
        df["Attempts"] = 0
    if "Notes" not in df.columns:
        df["Notes"] = ""

    # Coerce dtypes
    for dtcol in ["MeetingDateTime", "CallbackDateTime", "LastCallDateTime"]:
        try:
            df[dtcol] = pd.to_datetime(df[dtcol], errors='coerce')
        except Exception:
            pass
    try:
        df["Attempts"] = pd.to_numeric(df["Attempts"], errors='coerce').fillna(0).astype(int)
    except Exception:
        pass

    # Reorder columns
    df = df[[c for c in APP_COLUMNS if c in df.columns]]
    return df

def _example_template_subject():
    return "Quick intro – {name}"

def _example_template_body():
    return (
        "Hi {first_name},\n\n"
        "Great speaking with you. Confirming our meeting on {meeting_date} at {meeting_time}.\n\n"
        "If that time changes, just reply to this email.\n\n"
        "Best,\nYour Name"
    )

def _first_name(full_name: str) -> str:
    full_name = (full_name or "").strip()
    return full_name.split(" ")[0] if full_name else ""

def _fmt_date(dt) -> str:
    if pd.isna(dt) or dt is None or dt == "":
        return ""
    if isinstance(dt, str):
        try:
            dt = pd.to_datetime(dt)
        except Exception:
            return dt
    return dt.strftime("%B %d, %Y")

def _fmt_time(dt) -> str:
    if pd.isna(dt) or dt is None or dt == "":
        return ""
    if isinstance(dt, str):
        try:
            dt = pd.to_datetime(dt)
        except Exception:
            return dt
    return dt.strftime("%I:%M %p").lstrip('0')

def _render_template(row, subject_tmpl: str, body_tmpl: str):
    name = row.get("Name", "")
    meeting_dt = row.get("MeetingDateTime", None)
    mapping = {
        "name": name,
        "first_name": _first_name(name),
        "company": row.get("Company", ""),
        "meeting_datetime": "" if pd.isna(meeting_dt) else str(meeting_dt),
        "meeting_date": _fmt_date(meeting_dt),
        "meeting_time": _fmt_time(meeting_dt),
    }
    try:
        subject = subject_tmpl.format(**mapping)
    except Exception:
        subject = subject_tmpl
    try:
        body = body_tmpl.format(**mapping)
    except Exception:
        body = body_tmpl
    return subject, body

def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def _to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

def _touch_datetime(row):
    # Prefer LastCallDateTime, then Callback, then Meeting
    for c in ["LastCallDateTime", "CallbackDateTime", "MeetingDateTime"]:
        dt = row.get(c, pd.NaT)
        if not pd.isna(dt):
            return dt
    return pd.NaT

# ----------------------------- Sidebar -----------------------------
st.sidebar.header("📤 Import / Export")

uploaded = st.sidebar.file_uploader(
    "Upload prospects Excel (.xlsx or .xls)", type=["xlsx", "xls"]
)

if "df" not in st.session_state:
    st.session_state.df = _standardize_columns(pd.DataFrame(columns=REQUIRED_COLUMNS + OPTIONAL_COLUMNS))

if uploaded is not None:
    try:
        df_in = pd.read_excel(uploaded, engine="openpyxl")
    except Exception:
        df_in = pd.read_excel(uploaded, engine="xlrd")
    st.session_state.df = _standardize_columns(df_in)

# Calendar defaults
st.sidebar.header("🗓️ Calendar Defaults")
tz = st.sidebar.selectbox("Your timezone", COMMON_TIMEZONES, index=0)
default_duration = st.sidebar.number_input("Default meeting duration (minutes)", min_value=15, max_value=240, value=30, step=5)
organizer_name = st.sidebar.text_input("Organizer name", value="Your Name")
organizer_email = st.sidebar.text_input("Organizer email", value="you@example.com")
default_location = st.sidebar.text_input("Location (e.g., Zoom/Phone)", value="Phone")
desc_tmpl = st.sidebar.text_area("Invite description template", value=(
    "Meeting with {name} ({company}).\n\n"
    "Notes: {notes}\n"
), help="Placeholders: {name}, {company}, {notes}")

# ----------------------------- Main: Data Table -----------------------------
st.title("📇 Prospecting Manager")
st.caption("Import from Excel, track call outcomes and attempts, generate calendar invites and personalized emails, and analyze performance.")

df = st.session_state.df.copy()

with st.expander("Filters", expanded=False):
    status_filter = st.multiselect("Status", STATUS_OPTIONS[1:], default=[])
    state_filter = st.multiselect("State", sorted([s for s in df["State"].dropna().unique() if str(s).strip()]), default=[])
    company_filter = st.multiselect("Company", sorted([s for s in df["Company"].dropna().unique() if str(s).strip()]), default=[])
    df_view = df.copy()
    if status_filter:
        df_view = df_view[df_view["Status"].isin(status_filter)]
    if state_filter:
        df_view = df_view[df_view["State"].isin(state_filter)]
    if company_filter:
        df_view = df_view[df_view["Company"].isin(company_filter)]

st.write("### Prospect List")

edited = st.data_editor(
    df_view,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Status": st.column_config.SelectboxColumn(
            label="Status",
            options=STATUS_OPTIONS,
            required=False,
            help="Outcome of the last touch"
        ),
        "MeetingDateTime": st.column_config.DatetimeColumn(
            label="Meeting Date/Time",
            format="YYYY-MM-DD HH:mm"
        ),
        "CallbackDateTime": st.column_config.DatetimeColumn(
            label="Callback Date/Time",
            format="YYYY-MM-DD HH:mm"
        ),
        "LastCallDateTime": st.column_config.DatetimeColumn(
            label="Last Call Date/Time",
            format="YYYY-MM-DD HH:mm"
        ),
        "Attempts": st.column_config.NumberColumn(
            label="Attempts",
            min_value=0,
            step=1
        ),
        "Notes": st.column_config.TextColumn(
            label="Notes",
            help="Any notes about the conversation"
        ),
    },
    hide_index=True,
)

# Merge edits back into session_state.df using index alignment
if not df_view.empty:
    mask = df.index.isin(df_view.index)
    st.session_state.df.loc[mask, :] = edited.values
else:
    st.session_state.df = _standardize_columns(edited)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.download_button(
        label="⬇️ Download updated dataset (Excel)",
        data=_to_excel_bytes(st.session_state.df),
        file_name="prospects_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col2:
    contacts_cols = ["Name", "Phone", "Email"]
    st.download_button(
        label="⬇️ Export contacts (Name/Phone/Email – CSV)",
        data=_to_csv_bytes(st.session_state.df[contacts_cols]),
        file_name="contacts_export.csv",
        mime="text/csv",
        use_container_width=True,
    )
with col3:
    if st.button("➕ Increment attempts for filtered view"):
        st.session_state.df.loc[df_view.index, "Attempts"] = st.session_state.df.loc[df_view.index, "Attempts"].astype(int) + 1
        st.toast(f"Incremented attempts for {len(df_view)} prospect(s).", icon="✅")
with col4:
    st.success(f"Total prospects: {len(st.session_state.df)}")

# ----------------------------- Calendar Invites -----------------------------
st.write("---")
st.subheader("🗓️ Calendar Invites (.ics)")

def _ics_event_bytes(row, tzname: str, duration_min: int, organizer_name: str, organizer_email: str, location: str, desc_template: str) -> bytes:
    name = str(row.get("Name", "")).strip()
    company = str(row.get("Company", "")).strip()
    notes = str(row.get("Notes", "")).strip()
    meeting_dt = row.get("MeetingDateTime", pd.NaT)
    if pd.isna(meeting_dt):
        return b""
    # Treat naive as local tz
    if meeting_dt.tzinfo is None:
        local = ZoneInfo(tzname)
        start_local = meeting_dt.replace(tzinfo=local)
    else:
        start_local = meeting_dt
    end_local = start_local + timedelta(minutes=int(duration_min))
    # Convert to UTC for ICS
    start_utc = start_local.astimezone(ZoneInfo("UTC"))
    end_utc = end_local.astimezone(ZoneInfo("UTC"))
    uid = f"{uuid.uuid4()}@prospecting-app"
    now_utc = datetime.utcnow().replace(tzinfo=ZoneInfo("UTC"))
    summary = f"Meeting – {name}" if company == "" else f"Meeting – {name} ({company})"
    description = desc_template.format(name=name, company=company, notes=notes)
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Prospecting Manager//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{now_utc.strftime('%Y%m%dT%H%M%SZ')}",
        f"DTSTART:{start_utc.strftime('%Y%m%dT%H%M%SZ')}",
        f"DTEND:{end_utc.strftime('%Y%m%dT%H%M%SZ')}",
        f"SUMMARY:{summary}",
        f"DESCRIPTION:{description.replace('\n', '\\n')}",
        f"LOCATION:{location}",
        f"ORGANIZER;CN={organizer_name}:mailto:{organizer_email}",
    ]
    email = str(row.get("Email", "")).strip()
    if email:
        lines.append(f"ATTENDEE;CN={name};ROLE=REQ-PARTICIPANT:mailto:{email}")
    lines.extend(["END:VEVENT", "END:VCALENDAR"])
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")

yes_mask = (st.session_state.df["Status"].astype(str).str.lower() == "yes") & (~st.session_state.df["MeetingDateTime"].isna())
yes_df = st.session_state.df.loc[yes_mask]

colA, colB = st.columns([2, 1])
with colA:
    st.write(f"**Meetings to generate invites for:** {len(yes_df)}")
    if len(yes_df):
        st.dataframe(yes_df[["Name", "Company", "Email", "MeetingDateTime", "Notes"]], use_container_width=True)
with colB:
    if len(yes_df):
        events = []
        for _, r in yes_df.iterrows():
            ics_bytes = _ics_event_bytes(r, tz, default_duration, organizer_name, organizer_email, default_location, desc_tmpl)
            if ics_bytes:
                events.append(ics_bytes.decode("utf-8"))
        if events:
            bulk = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Prospecting Manager//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
            for e in events:
                inner = e.split("BEGIN:VEVENT")
                if len(inner) > 1:
                    bulk.append("BEGIN:VEVENT" + inner[1].split("END:VCALENDAR")[0].split("END:VEVENT")[0] + "END:VEVENT")
            bulk.append("END:VCALENDAR")
            bulk_bytes = ("\r\n".join(bulk) + "\r\n").encode("utf-8")
            st.download_button(
                label="⬇️ Download all meeting invites (.ics)",
                data=bulk_bytes,
                file_name="meetings_bulk.ics",
                mime="text/calendar",
                use_container_width=True,
            )
    st.caption("Tip: You can also create individual invites below.")

if len(yes_df):
    st.write("**Individual invites**")
    for idx, row in yes_df.iterrows():
        filename = f"{row.get('Name','meeting').replace(' ', '_')}_{pd.to_datetime(row['MeetingDateTime']).strftime('%Y%m%dT%H%M')}.ics"
        ics = _ics_event_bytes(row, tz, default_duration, organizer_name, organizer_email, default_location, desc_tmpl)
        if ics:
            st.download_button(
                label=f"Download invite: {row.get('Name','')}",
                data=ics,
                file_name=filename,
                mime="text/calendar",
            )

# ----------------------------- Email Templates -----------------------------
st.write("---")
st.subheader("✉️ Email Template")

with st.container(border=True):
    subject_tmpl = st.text_input("Subject template", value=_example_template_subject(),
                                 help="Use placeholders: {name}, {first_name}, {company}, {meeting_date}, {meeting_time}")
    body_tmpl = st.text_area(
        "Body template",
        value=_example_template_body(),
        height=200,
        help="Use placeholders: {name}, {first_name}, {company}, {meeting_datetime}, {meeting_date}, {meeting_time}"
    )

    if not st.session_state.df.empty:
        idx = st.selectbox(
            "Preview for",
            options=st.session_state.df.index.tolist(),
            format_func=lambda i: f"{st.session_state.df.loc[i, 'Name']} <{st.session_state.df.loc[i, 'Email']}>"
        )
        row = st.session_state.df.loc[idx]
        subject, body = _render_template(row, subject_tmpl, body_tmpl)

        st.write("**Preview – Subject:**", subject)
        st.code(body)

        email = str(row.get("Email", "")).strip()
        if email:
            mailto = f"mailto:{email}?subject={quote(subject)}&body={quote(body)}"
            st.markdown(f"[Open in your email client]({mailto})")

    if st.button("Generate personalized emails CSV for all with emails"):
        records = []
        for _, r in st.session_state.df.iterrows():
            email = str(r.get("Email", "")).strip()
            if not email:
                continue
            subject, body = _render_template(r, subject_tmpl, body_tmpl)
            records.append({
                "Name": r.get("Name", ""),
                "Email": email,
                "Subject": subject,
                "Body": body,
            })
        if records:
            emails_df = pd.DataFrame(records)
            st.download_button(
                label="⬇️ Download personalized_emails.csv",
                data=emails_df.to_csv(index=False).encode("utf-8"),
                file_name="personalized_emails.csv",
                mime="text/csv",
            )
        else:
            st.info("No rows with email addresses to generate.")

# ----------------------------- Analytics -----------------------------
st.write("---")
st.subheader("📈 Analytics – Success Rates & Patterns")

dfA = st.session_state.df.copy()
dfA["Attempted"] = (dfA["Attempts"].fillna(0) > 0) | (dfA["Status"].astype(str).str.len() > 0)
dfA["IsYes"] = dfA["Status"].astype(str).str.lower().eq("yes")

# Overall
total_attempted = int(dfA["Attempted"].sum())
total_yes = int(dfA["IsYes"].sum())
overall_rate = (total_yes / total_attempted * 100.0) if total_attempted else 0.0

c1, c2, c3 = st.columns(3)
c1.metric("Attempted", total_attempted)
c2.metric("Yes", total_yes)
c3.metric("Yes Rate", f"{overall_rate:.1f}%")

def rate_table(group_col: str, top_n: int = 15):
    tmp = dfA.copy()
    tmp[group_col] = tmp[group_col].fillna("").replace({None: ""})
    grp = tmp.groupby(group_col).agg(
        n=(group_col, 'size'),
        attempted=("Attempted", 'sum'),
        yes=("IsYes", 'sum')
    ).reset_index()
    grp = grp[grp['attempted'] > 0]
    grp["yes_rate"] = (grp["yes"] / grp["attempted"] * 100).round(1)
    grp = grp.sort_values(["yes_rate", "attempted"], ascending=[False, False]).head(top_n)
    return grp

st.write("#### By Company")
comp_tbl = rate_table("Company")
st.dataframe(comp_tbl, use_container_width=True)

st.write("#### By State")
state_tbl = rate_table("State")
st.dataframe(state_tbl, use_container_width=True)

# Time-of-day / Weekday from LastCallDateTime (or fallback)
touch_dt = dfA.apply(_touch_datetime, axis=1)
dfA["TouchDT"] = pd.to_datetime(touch_dt, errors='coerce')
dfA_valid = dfA.dropna(subset=["TouchDT"]).copy()
if not dfA_valid.empty:
    dfA_valid["Hour"] = dfA_valid["TouchDT"].dt.hour
    dfA_valid["Weekday"] = dfA_valid["TouchDT"].dt.day_name()

    st.write("#### Yes Rate by Hour of Day (based on Last Call/Touch)")
    hour_grp = dfA_valid.groupby("Hour").agg(
        attempted=("Attempted", 'sum'), yes=("IsYes", 'sum')
    )
    hour_grp = hour_grp[hour_grp["attempted"] > 0]
    hour_grp["yes_rate"] = (hour_grp["yes"] / hour_grp["attempted"] * 100).round(1)
    st.bar_chart(hour_grp["yes_rate"], use_container_width=True)

    st.write("#### Yes Rate by Weekday")
    wk_grp = dfA_valid.groupby("Weekday").agg(
        attempted=("Attempted", 'sum'), yes=("IsYes", 'sum')
    )
    wk_grp = wk_grp[wk_grp["attempted"] > 0]
    wk_grp["yes_rate"] = (wk_grp["yes"] / wk_grp["attempted"] * 100).round(1)
    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    wk_grp = wk_grp.reindex(weekday_order).dropna(how='all')
    st.bar_chart(wk_grp["yes_rate"], use_container_width=True)
else:
    st.info("Add values to **Last Call Date/Time** to analyze by hour or weekday.")

st.write("---")
st.caption("Note: Handle contact data responsibly and comply with your organization's privacy & outreach policies.")
