"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected one month at a
time via the in-page "Reporting Month" dropdown. See discover_monthly_files().

Sheets used (VRM)
─────────────────
  Active VT            — primary operational data (one row per vol-offering)
  Dropped VT            — volunteers who dropped their offering(s)
  Newly Registered VT  — all volunteers who registered in the current period

Sheets used (DRM)
─────────────────
  SESSION DUMP         — session-level log (1 row per session)
  DRM                  — centre-level summary
  DRM1                 — alternate centre summary with total attendance
  Offering Details     — one row per offering (grade × subject × centre)
  Active centers       — centre master with registration data
  New Enrolled student — mid-month new enrolments

Column mapping (VRM new file → dashboard internal name) 
────────────────────────────────────────────────────────
  Volunteer ID              → vol_id
  Volunteer Name            → vol_name
  EV Joined Date            → joined_dt
  Volunteer State           → res_state
  Volunteer City            → res_city
  Center Name               → center
  Center State              → state
  Offering Status           → status
  Total Hours (Comp+Offline)→ vol_hrs
  Attendance %              → attendance  (already float in this file)
  En Boys / En Girls        → gender split (NEW — not in old CSV)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
# Monthly data files live alongside this script and follow the naming
# convention VRM_<Month>_<Year>.xlsx / DRM_<Month>_<Year>.xlsx, e.g.
# "VRM_April_2026.xlsx" / "DRM_April_2026.xlsx". See discover_monthly_files().
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
# (case-insensitive, tolerant of one-or-more underscores between tokens)
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

# Sheets (VRM) — target/canonical sheet names. Actual lookup is done via
# _find_sheet(), which resolves these case/whitespace-insensitively so that
# month-to-month export quirks (e.g. "Active Centers" vs "Active centers")
# don't break loading.
SHEET_ACTIVE  = "Active VT"
SHEET_DROPPED = "Dropped VT"
SHEET_NEW_REG = "Newly Registered VT "   # trailing space is in the file


def _month_sort_key(month_name: str) -> int:
    """Map a month name ('April', 'Apr', ...) to its calendar number (1-12)
    for chronological sorting. Unrecognised names sort first (0)."""
    for fmt in ("%B", "%b"):
        try:
            return datetime.strptime(month_name.strip(), fmt).month
        except ValueError:
            continue
    return 0


def discover_monthly_files(data_dir: str) -> dict:
    """
    Scan `data_dir` for files matching the VRM_<Month>_<Year>.xlsx /
    DRM_<Month>_<Year>.xlsx naming convention (case-insensitive) and group
    them by reporting month. Any other files in the directory are ignored.

    Returns an ordered dict (oldest → newest), e.g.:
        {
          "April 2026": {"vrm": "/path/VRM_April_2026.xlsx", "drm": "/path/DRM_April_2026.xlsx"},
          "May 2026":   {"vrm": "/path/VRM_May_2026.xlsx",   "drm": None},
        }
    A month is only included if a VRM file was found for it — the DRM file
    is optional per month, since the DRM Client Report tab already shows
    its own "file not found" message when missing.
    """
    months: dict[str, dict] = {}
    try:
        filenames = os.listdir(data_dir)
    except OSError:
        filenames = []

    for fname in filenames:
        match = MONTH_FILE_RE.match(fname)
        if not match:
            continue
        kind, month_name, year = match.group(1).upper(), match.group(2).title(), match.group(3)
        label = f"{month_name} {year}"
        entry = months.setdefault(label, {"vrm": None, "drm": None})
        entry["vrm" if kind == "VRM" else "drm"] = os.path.join(data_dir, fname)

    months = {label: paths for label, paths in months.items() if paths["vrm"]}

    def _sort_key(label: str):
        month_name, year = label.rsplit(" ", 1)
        return (int(year), _month_sort_key(month_name))

    return dict(sorted(months.items(), key=lambda kv: _sort_key(kv[0])))


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    """
    Resolve a sheet name case/whitespace-insensitively against the sheets
    actually present in `xl` (different months' exports sometimes vary
    casing/spacing, e.g. "Active centers" vs "Active Centers"). Falls back
    to the literal `target` — and lets pandas raise its normal error — if
    no close match is found.
    """
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
    """
    Rename legacy/alternate column names to their canonical name, but only
    when the canonical column isn't already present. `aliases` maps
    canonical_name -> [alt_name_1, alt_name_2, ...]. Used to smooth over
    month-to-month column-naming drift in source workbooks.
    """
    rename_map = {}
    for canonical, alternates in aliases.items():
        if canonical in df.columns:
            continue
        for alt in alternates:
            if alt in df.columns:
                rename_map[alt] = canonical
                break
    return df.rename(columns=rename_map) if rename_map else df

# ── Colour palette ────────────────────────────────────────────────────────────
P = {
    "teal":    "#0094c9",
    "green":   "#00964d",
    "orange":  "#f27c48",
    "red":     "#ed1c2d",
    "violet":  "#6c5ce7",
    "amber":   "#fdcb6e",
    "sky":     "#74b9ff",
    "mint":    "#00b894",
    "coral":   "#e17055",
    "lavender":"#a29bfe",
    "salmon":  "#fab1a0",
    "aqua":    "#55efc4",
}
SEQ  = list(P.values())
BG   = "rgba(0,0,0,0)"
GRID = "#e9ecef"

# ── DRM accent colours (richer palette for client tab) ────────────────────────
D = {
    "teal":   "#0f8a6e",
    "blue":   "#185fa5",
    "amber":  "#ba7517",
    "coral":  "#993c1d",
    "purple": "#534ab7",
    "green":  "#3b6d11",
    "red":    "#a32d2d",
    "gray":   "#5f5e5a",
}

# ── Subject normalisation ─────────────────────────────────────────────────────
SUBJECT_MAP = {
    "Conceptual Learning - Math":     "Math",
    "Conceptual Learning Math-HM":    "Math",
    "Maths - Worksheet":              "Math",
    "Conceptual Learning - Science":  "Science",
    "Conceptual Learning Science-HM": "Science",
    "English":                        "English",
    "English - Worksheet":            "English",
    "Spoken English: Level 1":        "Spoken English",
    "Spoken English: Level 2":        "Spoken English",
    "Basic Digital Literacy":         "Digital Literacy",
    "Artificial Intelligence (AI)":   "AI",
    "Explore Coding":                 "Coding",
    "Concise content 1":              "Concise Content",
    "Concise Content 1":              "Concise Content",
    "Concise content 2":              "Concise Content",
    "Guest Sessions":                 "Guest Sessions",
    "Scholarship":                    "Scholarship",
    "Reading Program":                "Reading Program",
}

# ── Profession normalisation ──────────────────────────────────────────────────
PROFESSION_MAP = {
    "corporates":                              "Corporate Professional",
    "others":                                  "Others",
    "student_ug":                              "Student (UG)",
    "home_makers":                             "Home Maker",
    "Housewife":                               "Home Maker",
    "house wife":                              "Home Maker",
    "Home Maker":                              "Home Maker",
    "Home maker":                              "Home Maker",
    "teacher_school":                          "School Teacher",
    "teacher_university":                      "University Teacher",
    "Teaching":                                "School Teacher",
    "teaching":                                "School Teacher",
    "Educator":                                "School Teacher",
    "Principal of a Primary school":           "School Teacher",
    "retired teacher":                         "Retired Professional",
    "student_pg":                              "Student (PG)",
    "Student":                                 "Student (UG)",
    "Student - PG":                            "Student (PG)",
    "business/self_employed":                  "Business / Self-Employed",
    "Self-employed":                           "Business / Self-Employed",
    "self employed":                           "Business / Self-Employed",
    "Service":                                 "Business / Self-Employed",
    "retired_professional":                    "Retired Professional",
    "Retd. Professional":                      "Retired Professional",
    "top_managemenet":                         "Management",
    "Management":                              "Management",
    "finance":                                 "Finance",
    "legal":                                   "Legal",
    "Legal":                                   "Legal",
    "Education, Training, and Library":        "School Teacher",
    "Computer and Mathematical":               "Corporate Professional",
    "Business and Financial Operations":       "Finance",
    "Healthcare Practitioners and Technical":  "Healthcare",
    "Medical Professional":                    "Healthcare",
    "Life, Physical, and Social Science":      "Others",
    "Office and Administrative Support":       "Others",
    "Production/Manufacturing":                "Others",
    "Sales and Related":                       "Business / Self-Employed",
    "Community and Social Service":            "Others",
    "Farming, Fishing, and Forestry":          "Others",
    "Trainer, manufacturers":                  "Others",
}

# ── Reference channel bucketing ───────────────────────────────────────────────
def _group_ref(ref: str) -> str:
    r = str(ref).strip()
    if r in ("Internet Search", "Facebook", "Community Outreach"):
        return "Online / Direct"
    if r in ("Emailer", "DCP Campaign", "eVidyaloka", "eVidyaloka Trust"):
        return "eVidyaloka Campaign"
    if r in ("Word of Mouth",):
        return "Word of Mouth"
    if any(k in r for k in (
        "KPMG", "Infosys", "Tech Mahindra", "HPInc", "HPE", "L&T",
        "HSBC", "EY", "Adobe", "Brillio", "Broadridge", "Accenture",
        "CISCO", "Fidelity", "ConnectFor", "Microsoft", "Cognizant",
        "Firstsource", "Pricewaterhousecoopers", "Reliance Foundation",
        "Scaler", "Joy of Reading",
    )):
        return "Corporate / Partner Referral"
    if any(k in r for k in (
        "College", "University", "BMS", "NIT", "IIT", "IIM",
    )):
        return "Academic Institution"
    if any(k in r for k in (
        "Bhumi", "Udaan", "Swabhiman", "Foundation", "NGO",
    )):
        return "NGO / Community Org"
    return "Other"

# ─────────────────────────────────────────────────────────────────────────────
# VRM DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> dict:
    """
    Returns a dict with three clean DataFrames:
      active   — Active VT sheet, fully normalised
      dropped  — Dropped VT sheet
      new_reg  — Newly Registered VT sheet
    """
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    # ── Active VT ─────────────────────────────────────────────────────────────
    active = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_ACTIVE))
    active.columns = active.columns.str.strip()

    active.rename(columns={
        "Volunteer ID":               "Volunteer id",
        "Volunteer Name":             "Volunteer name",
        "EV Joined Date":             "Joined(ev)",
        "Volunteer State":            "Residence state",
        "Volunteer City":             "Residence city",
        "Center Name":                "Center name",
        "Center State":               "State",
        "Offering Status":            "Offering status",
        "Total Hours (Comp + Offline)":"Total hours(Comp+Offline)",
        "Attendance %":               "Attendance%",
    }, inplace=True)

    for col in ["State", "Residence state"]:
        if col in active.columns:
            active[col] = (
                active[col].astype(str)
                .str.replace("\u00ac\u00a0", " ", regex=False)
                .str.replace("\u00a0", " ", regex=False)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )

    active["Attendance%"] = pd.to_numeric(active["Attendance%"], errors="coerce")
    active["Joined(ev)"]  = pd.to_datetime(active["Joined(ev)"], errors="coerce")

    active["Subject_clean"]    = active["Subject"].map(SUBJECT_MAP).fillna(active["Subject"])
    active["Profession_clean"] = active["Profession"].map(PROFESSION_MAP).fillna("Others")
    active["Ref_group"]        = active["Reference"].apply(_group_ref)

    for col in ["Donor", "State", "Subject_clean", "Center name",
                "Residence state", "Residence city", "Ref_group"]:
        if col in active.columns:
            active[col] = active[col].fillna("Unknown")

    for col in ["Registered", "Reg Boys", "Reg Girls",
                "Enrolled",   "En Boys",  "En Girls",
                "Planned",    "Scheduled", "Completed",
                "Offline",    "Cancelled",
                "Total hours(Comp+Offline)", "CLH"]:
        if col in active.columns:
            active[col] = pd.to_numeric(active[col], errors="coerce").fillna(0).astype(int)

    # ── Dropped VT ────────────────────────────────────────────────────────────
    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))
    dropped.columns = dropped.columns.str.strip()
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    # ── Newly Registered VT ───────────────────────────────────────────────────
    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

# ─────────────────────────────────────────────────────────────────────────────
# TOPIC and SubTopic DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
TOPIC_FILE = os.path.join(DATA_DIR, "Topic_SubTopic_Cancelled_Offline_May_2026.xlsx")

@st.cache_data(show_spinner=False)
def load_topic_data():
    if not os.path.exists(TOPIC_FILE):
        return None, None, None
    topic_df = pd.read_excel(TOPIC_FILE, sheet_name="Topic-SubTopic")
    cancelled_df = pd.read_excel(TOPIC_FILE, sheet_name="Cancelled Sessions")
    offline_df = pd.read_excel(TOPIC_FILE, sheet_name="Offline Sessions")
    return topic_df, cancelled_df, offline_df

# ─────────────────────────────────────────────────────────────────────────────
# DRM DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
SUBJECT_NORM_DRM = {
    "Conceptual Learning - Math":     "Math",
    "Conceptual Learning Math-HM":    "Math",
    "Conceptual Learning - Science":  "Science",
    "Conceptual Learning Science-HM": "Science",
    "Concise Content 1":              "Concise Content",
    "Concise content 1":              "Concise Content",
    "English":                        "English",
    "Spoken English: Level 1":        "Spoken English",
    "Guest Sessions":                 "Guest Sessions",
}

# Some months' DRM exports use different column names on the "DRM"
# (centre-level summary) sheet. Map canonical name -> legacy/alternate
# names; _apply_aliases() renames the first match found, only when the
# canonical column isn't already present.
DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    """Load and clean all sheets from the DRM workbook."""
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

    # SESSION DUMP
    sess = pd.read_excel(xl, sheet_name=_find_sheet(xl, "SESSION DUMP"))
    sess.columns = sess.columns.str.strip()
    sess["Session_start"] = pd.to_datetime(sess["Session_start"], errors="coerce")
    sess["Attendance%"]   = pd.to_numeric(sess["Attendance%"], errors="coerce")
    sess["State"]         = sess["State"].apply(_fix_state)
    sess["Donor"]         = sess["Donor"].fillna("Unknown")
    sess["Subject_clean"] = sess["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(sess["Subject"])
    for col in ["Present/CLH", "Total students", "Boys", "Girls"]:
        if col in sess.columns:
            sess[col] = pd.to_numeric(sess[col], errors="coerce").fillna(0).astype(int)
    sess["week"] = sess["Session_start"].dt.to_period("W").astype(str)
    sess["dow"]  = sess["Session_start"].dt.day_name()
    sess["hour"] = sess["Session_start"].dt.hour

    # DRM (centre-level summary)
    drm = pd.read_excel(xl, sheet_name=_find_sheet(xl, "DRM"))
    drm.columns = drm.columns.str.strip()
    drm = _apply_aliases(drm, DRM_SHEET_ALIASES)
    drm["State"]      = drm["State"].apply(_fix_state)
    drm["Donor Name"] = drm["Donor Name"].fillna("Unknown")
    for col in ["Planned","Scheduled","Completed","Offline","Cancelled",
                "Live Volunteers","Live_CLH","Registered","Reg Boys","Reg Girls",
                "Enrolled","En Boys","En Girls"]:
        if col in drm.columns:
            drm[col] = pd.to_numeric(drm[col], errors="coerce").fillna(0)
    drm["Attendance %"]  = pd.to_numeric(drm["Attendance %"], errors="coerce")
    drm["Completion%"]   = ((drm["Completed"] + drm["Offline"])
                            / drm["Planned"].replace(0, float("nan")) * 100).round(1)
    drm["Cancellation%"] = (drm["Cancelled"]
                            / drm["Planned"].replace(0, float("nan")) * 100).round(1)
    drm["Dropout%"]      = ((drm["Registered"] - drm["Enrolled"])
                            / drm["Registered"].replace(0, float("nan")) * 100).round(1)
    drm["Girl%"]         = (drm["En Girls"]
                            / (drm["En Boys"] + drm["En Girls"]).replace(0, float("nan")) * 100).round(1)

    # ACTIVE CENTERS
    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))
    ac.columns = ac.columns.str.strip()

    # OFFERING DETAILS
    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

    # NEW ENROLLED STUDENT
    ne = pd.read_excel(xl, sheet_name=_find_sheet(xl, "New Enrolled student"))
    ne.columns = ne.columns.str.strip()

    return {"sess": sess, "drm": drm, "ac": ac, "od": od, "ne": ne}


# ─────────────────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _layout(fig, height=380, legend_bottom=False, margin=None):
    m = margin or dict(l=0, r=0, t=40, b=0)
    fig.update_layout(
        height=height,
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(family="Inter, Helvetica, sans-serif", size=12, color="#2d3436"),
        margin=m,
    )
    fig.update_xaxes(showgrid=False, linecolor="#dee2e6", linewidth=1)
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False)
    if legend_bottom:
        fig.update_layout(legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, title_text="",
        ))
    return fig


def _kpi(label: str, value: str, color: str, sub: str = "") -> str:
    return (
        f"<div style='background:#f8f9fa;border-radius:10px;padding:14px 10px;"
        f"border-left:4px solid {color};text-align:center;'>"
        f"<div style='font-size:0.72em;color:#636e72;font-weight:600;"
        f"letter-spacing:0.05em;text-transform:uppercase;margin-bottom:4px;'>{label}</div>"
        f"<div style='font-size:1.7rem;font-weight:800;color:{color};"
        f"line-height:1.15;'>{value}</div>"
        f"<div style='font-size:0.7em;color:#b2bec3;margin-top:3px;'>{sub}</div>"
        f"</div>"
    )


def _drm_kpi(label: str, value: str, color: str, sub: str = "") -> str:
    """Larger KPI card for the premium DRM client tab."""
    return (
        f"<div style='background:white;border-radius:12px;padding:18px 14px;"
        f"border-left:5px solid {color};text-align:center;"
        f"box-shadow:0 2px 8px rgba(0,0,0,0.06);'>"
        f"<div style='font-size:0.68em;color:#636e72;font-weight:700;"
        f"letter-spacing:0.07em;text-transform:uppercase;margin-bottom:5px;'>{label}</div>"
        f"<div style='font-size:1.9rem;font-weight:800;color:{color};"
        f"line-height:1.1;font-family:\"DM Mono\",monospace;'>{value}</div>"
        f"<div style='font-size:0.68em;color:#b2bec3;margin-top:4px;'>{sub}</div>"
        f"</div>"
    )


def _insight_box(text: str, color: str = "#ba7517", bg: str = "#faeeda") -> str:
    """Inline insight callout box."""
    return (
        f"<div style='background:{bg};border-left:4px solid {color};"
        f"border-radius:8px;padding:12px 16px;margin-top:12px;'>"
        f"<span style='color:{color};font-weight:700;font-size:0.8em;'>KEY INSIGHT&nbsp;&nbsp;</span>"
        f"<span style='color:#2d3436;font-size:0.82em;'>{text}</span>"
        f"</div>"
    )


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RENDER
# ─────────────────────────────────────────────────────────────────────────────
def render_ops_dashboard():
    st.title("🏢 Operations & Impact Command Center")

    # ── Month selector — auto-detects VRM_<Month>_<Year> / DRM_<Month>_<Year> files ──
    available_months = discover_monthly_files(DATA_DIR)

    if not available_months:
        st.markdown(
            "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
            "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health</p>",
            unsafe_allow_html=True,
        )
        st.markdown("---")
        st.error(
            "⚠️ No monthly data files found. Place files named like "
            "`VRM_April_2026.xlsx` (and optionally `DRM_April_2026.xlsx`) "
            f"in `{DATA_DIR}`."
        )
        return

    month_labels = list(available_months.keys())
    selected_month = st.selectbox(
        "📅 Reporting Month",
        options=month_labels,
        index=len(month_labels) - 1,   # default to most recent month
        key="selected_month",
    )
    vrm_path = available_months[selected_month]["vrm"]
    drm_path = available_months[selected_month]["drm"]

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {selected_month}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Load VRM data ─────────────────────────────────────────────────────────
    with st.spinner("Loading VRM dataset…"):
        data = load_data(vrm_path)

    if not data:
        st.error(
            f"⚠️ Could not load the VRM file for {selected_month} "
            f"(`{os.path.basename(vrm_path)}`)."
        )
        return

    df_raw  = data["active"]
    dropped = data["dropped"]
    new_reg = data["new_reg"]

    # ── Active-tab tracking via session_state ────────────────────────────────
    # Streamlit doesn't expose which tab is currently open, so we use two
    # small toggle buttons in the sidebar to let the user switch filter context.
    # This controls which filter group is rendered in the sidebar.
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ── Load DRM raw data upfront (cached) so sidebar can read its option lists
    with st.spinner("Loading DRM data…"):
        drm_data_raw = load_drm_data(drm_path) if drm_path else {}

    # ── Build sidebar — content switches based on active_tab flag ─────────────
    with st.sidebar:
        st.markdown("---")

        # Two toggle buttons to switch between VRM and DRM filter sets.
        # Clicking a button sets session_state.active_tab and triggers a rerun.
        tog1, tog2 = st.columns(2)
        if tog1.button("📋 VRM Filters", use_container_width=True,
                       type="primary" if st.session_state.active_tab == "vrm" else "secondary"):
            st.session_state.active_tab = "vrm"
            st.rerun()
        if tog2.button("📊 DRM Filters", use_container_width=True,
                       type="primary" if st.session_state.active_tab == "drm" else "secondary"):
            st.session_state.active_tab = "drm"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        # ── VRM FILTERS ───────────────────────────────────────────────────────
        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())
            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

            # Searchable multiselect: the placeholder text and compact widget
            # give a "search" feel — user types and the list filters live.
            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

            # Cascade: states narrow to only those linked to selected donors
            if sel_donors:
                _linked_states = sorted(
                    df_raw[df_raw["Donor"].isin(sel_donors)]["State"].dropna().unique()
                )
            else:
                _linked_states = _vrm_all_states

            sel_states = st.multiselect(
                "State (Centre)",
                options=_linked_states,
                default=_linked_states,
                placeholder="Type to search states…",
                key="ops_state",
            )

            # Cascade: subjects narrow to those in the selected donor + state combination
            _vmask = pd.Series([True] * len(df_raw), index=df_raw.index)
            if sel_donors:
                _vmask &= df_raw["Donor"].isin(sel_donors)
            if sel_states:
                _vmask &= df_raw["State"].isin(sel_states)
            _linked_subjects = sorted(df_raw[_vmask]["Subject_clean"].dropna().unique())

            sel_subjects = st.multiselect(
                "Subject",
                options=_linked_subjects,
                default=_linked_subjects,
                placeholder="Type to search subjects…",
                key="ops_subject",
            )

            st.caption(
                "Filters cascade — selecting a donor narrows available states & subjects. "
                "Applies to Volunteers, Centres and Academic Health tabs."
            )

            # DRM filter placeholders (full universe, unused while on VRM mode)
            if drm_data_raw:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        # ── DRM FILTERS ───────────────────────────────────────────────────────
        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw:
                _sr = drm_data_raw["sess"]
                _drm_all_donors   = sorted(_sr["Donor"].dropna().unique())
                _drm_all_states   = sorted(_sr["State"].dropna().unique())
                _drm_all_subjects = sorted(_sr["Subject_clean"].dropna().unique())

                sel_drm_donors = st.multiselect(
                    "Donor",
                    options=_drm_all_donors,
                    default=_drm_all_donors,
                    placeholder="Type to search donors…",
                    key="drm_donor",
                )

                # Cascade: states narrow to those linked to selected DRM donors
                if sel_drm_donors:
                    _drm_linked_states = sorted(
                        _sr[_sr["Donor"].isin(sel_drm_donors)]["State"].dropna().unique()
                    )
                else:
                    _drm_linked_states = _drm_all_states

                sel_drm_states = st.multiselect(
                    "State",
                    options=_drm_linked_states,
                    default=_drm_linked_states,
                    placeholder="Type to search states…",
                    key="drm_state",
                )

                # Cascade: subjects narrow to selected donor + state
                _dmask = pd.Series([True] * len(_sr), index=_sr.index)
                if sel_drm_donors:
                    _dmask &= _sr["Donor"].isin(sel_drm_donors)
                if sel_drm_states:
                    _dmask &= _sr["State"].isin(sel_drm_states)
                _drm_linked_subjects = sorted(_sr[_dmask]["Subject_clean"].dropna().unique())

                sel_drm_subjects = st.multiselect(
                    "Subject",
                    options=_drm_linked_subjects,
                    default=_drm_linked_subjects,
                    placeholder="Type to search subjects…",
                    key="drm_subject",
                )

                st.caption(
                    "Filters cascade — selecting a donor narrows available states & subjects. "
                    "Applies to the DRM Client Report tab only."
                )
            else:
                st.warning("DRM file not found — filters unavailable.")
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

            # VRM filter placeholders (full universe, unused while on DRM mode)
            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

    # ── Apply VRM filters to produce the working DataFrame ────────────────────
    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    vol_df = df.drop_duplicates(subset="Volunteer id").copy() if not df.empty else df.copy()

    # ─────────────────────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────────────────────
    tab_vol, tab_ctr, tab_aca, tab_drm, tab_top = st.tabs([
        "🙋 Volunteers",
        "🏫 Centres",
        "📚 Academic Health",
        "📊 DRM Client Report",
        "📚 Topic & SubTopic Analytics",
    ])


    # ═════════════════════════════════════════════════════════════════════════
    # TAB 1 — VOLUNTEERS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_vol:

        if df.empty:
            st.info(
                "🔍 No data available for the selected filters.\n\n"
                "Try broadening your Donor, State, or Subject selection in the sidebar."
            )
            st.stop()

        total_vols_offering = len(df)
        unique_vols         = vol_df["Volunteer id"].nunique()
        dropped_vols        = len(dropped)
        new_vols            = new_reg["User ID"].nunique()
        active_centres      = df["Center name"].nunique()
        total_enrolled      = int(df.drop_duplicates("Center name")["Enrolled"].sum())
        total_vol_hrs       = int(df["Total hours(Comp+Offline)"].sum())
        total_clh           = int(df["CLH"].sum())
        avg_att             = df["Attendance%"].mean()
        avg_att_display     = f"{avg_att:.1f}%" if pd.notna(avg_att) else "N/A"
        completion_rt       = (
            df["Completed"].sum() / df["Planned"].sum() * 100
            if df["Planned"].sum() > 0 else 0
        )

        st.markdown("#### Volunteer Overview")
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.markdown(_kpi("Total Volunteers",   f"{total_vols_offering:,}", P["teal"],   "vol-offering rows"),   unsafe_allow_html=True)
        kc2.markdown(_kpi("Unique Volunteers",  f"{unique_vols:,}",         P["green"],  "deduplicated"),        unsafe_allow_html=True)
        kc3.markdown(_kpi("Newly Registered",   f"{new_vols:,}",            P["violet"], "from registrations sheet"), unsafe_allow_html=True)
        kc4.markdown(_kpi("Dropped Volunteers", f"{dropped_vols:,}",        P["coral"],  "from dropped sheet"),  unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        ki1, ki2, ki3, ki4, ki5, ki6 = st.columns(6)
        ki1.markdown(_kpi("Active Centres",   f"{active_centres:,}",  P["teal"],   "unique"),             unsafe_allow_html=True)
        ki2.markdown(_kpi("Enrolled Students",f"{total_enrolled:,}",  P["green"],  "total seats"),        unsafe_allow_html=True)
        ki3.markdown(_kpi("Total Vol Hrs",    f"{total_vol_hrs:,}",   P["orange"], "Comp + Offline hrs"), unsafe_allow_html=True)
        ki4.markdown(_kpi("Total CLH",        f"{total_clh:,}",       P["violet"], "child learning hrs"), unsafe_allow_html=True)
        ki5.markdown(_kpi("Avg Attendance",   avg_att_display,        P["amber"],  "across sessions"),    unsafe_allow_html=True)
        ki6.markdown(_kpi("Class Completion", f"{completion_rt:.1f}%",P["mint"],   "completed / planned"),unsafe_allow_html=True)

        st.markdown("---")

        st.markdown("#### Volunteer Distribution")
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            st.markdown("##### By Reference / Acquisition Channel")
            ref_data = (
                vol_df.groupby("Ref_group")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=True)
            )
            fig_ref = px.bar(
                ref_data, x="Volunteers", y="Ref_group",
                orientation="h", text="Volunteers",
                color="Ref_group", color_discrete_sequence=SEQ,
            )
            fig_ref.update_traces(
                textposition="outside", showlegend=False,
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
            )
            _layout(fig_ref, height=320)
            fig_ref.update_layout(xaxis_title="Volunteers", yaxis_title="", showlegend=False)
            fig_ref.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_ref.update_yaxes(showgrid=False)
            st.plotly_chart(fig_ref, use_container_width=True)

        with col_b2:
            st.markdown("##### By Profession")
            prof_data = (
                vol_df.groupby("Profession_clean")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=False)
            )
            fig_prof = px.pie(
                prof_data, names="Profession_clean", values="Volunteers",
                hole=0.45, color_discrete_sequence=SEQ,
            )
            fig_prof.update_traces(
                textposition="outside",
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                pull=[0.02] * len(prof_data),
            )
            fig_prof.update_layout(
                height=320, showlegend=False,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=10, r=10, t=40, b=10),
                annotations=[dict(
                    text=f"<b>{unique_vols:,}</b><br>Vols",
                    x=0.5, y=0.5, font_size=15, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_prof, use_container_width=True)

        st.markdown("---")

        st.markdown("#### Volunteer Residence — by State")
        res_state = (
            vol_df.groupby("Residence state")["Volunteer id"].nunique()
            .reset_index()
            .rename(columns={"Volunteer id": "Volunteers"})
            .sort_values("Volunteers", ascending=True)
        )
        fig_res = px.bar(
            res_state, x="Volunteers", y="Residence state",
            orientation="h", text="Volunteers",
            color="Volunteers",
            color_continuous_scale=[P["sky"], P["teal"]],
        )
        fig_res.update_traces(
            textposition="outside", marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
        )
        fig_res.update_coloraxes(showscale=False)
        _layout(fig_res, height=max(300, len(res_state) * 28),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_res.update_layout(xaxis_title="Volunteers", yaxis_title="")
        fig_res.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_res.update_yaxes(showgrid=False)
        st.plotly_chart(fig_res, use_container_width=True)

        st.markdown("---")

        st.markdown("#### New Volunteer Registrations — Monthly Trend")
        st.caption("Source: Newly Registered VT sheet — all registrations regardless of sidebar filters.")
        trend_df = new_reg.dropna(subset=["Date Joined"]).copy()
        trend_df["YM"] = trend_df["Date Joined"].dt.to_period("M").astype(str)
        monthly = (
            trend_df.groupby("YM")["User ID"].nunique()
            .reset_index()
            .rename(columns={"User ID": "New Vols"})
            .sort_values("YM")
        )
        fig_trend = px.area(
            monthly, x="YM", y="New Vols",
            markers=True, color_discrete_sequence=[P["teal"]],
        )
        fig_trend.update_traces(
            line=dict(width=2.5),
            marker=dict(size=6, color=P["teal"]),
            fillcolor="rgba(0,148,201,0.12)",
            hovertemplate="<b>%{x}</b><br>New Vols: %{y:,}<extra></extra>",
        )
        _layout(fig_trend, height=240, margin=dict(l=0, r=0, t=20, b=60))
        fig_trend.update_layout(xaxis_title="", yaxis_title="New Volunteers")
        fig_trend.update_xaxes(tickangle=-45, showgrid=False)
        st.plotly_chart(fig_trend, use_container_width=True)

        st.markdown("---")
        st.markdown("#### Newly Registered Volunteers — Gender Split")
        if "Gender" in new_reg.columns:
            gen_new = (
                new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                .groupby("Gender")["User ID"].nunique()
                .reset_index()
                .rename(columns={"User ID": "Volunteers"})
            )
            col_g1, col_g2 = st.columns([1, 2])
            with col_g1:
                fig_gen = px.pie(
                    gen_new, names="Gender", values="Volunteers",
                    hole=0.5,
                    color="Gender",
                    color_discrete_map={"Male": P["teal"], "Female": P["coral"]},
                )
                fig_gen.update_traces(
                    texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                    textposition="outside",
                    hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                )
                fig_gen.update_layout(
                    height=280, showlegend=False,
                    plot_bgcolor=BG, paper_bgcolor=BG,
                    margin=dict(l=10, r=10, t=30, b=10),
                )
                st.plotly_chart(fig_gen, use_container_width=True)
            with col_g2:
                gen_ref = (
                    new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                    .assign(Ref_group=lambda x: x["Reference"].apply(_group_ref))
                    .groupby(["Ref_group", "Gender"])["User ID"].nunique()
                    .reset_index()
                    .rename(columns={"User ID": "Volunteers"})
                )
                fig_gen_ref = px.bar(
                    gen_ref, x="Volunteers", y="Ref_group", color="Gender",
                    orientation="h", barmode="stack",
                    color_discrete_map={"Male": P["teal"], "Female": P["coral"]},
                    text="Volunteers",
                )
                fig_gen_ref.update_traces(
                    textposition="inside", texttemplate="%{text}",
                    hovertemplate="<b>%{y}</b> · %{fullData.name}: %{x}<extra></extra>",
                )
                _layout(fig_gen_ref, height=280, legend_bottom=False,
                        margin=dict(l=0, r=0, t=30, b=0))
                fig_gen_ref.update_layout(
                    xaxis_title="Volunteers", yaxis_title="",
                    legend=dict(orientation="h", yanchor="top", y=-0.15,
                                xanchor="center", x=0.5, title_text=""),
                )
                fig_gen_ref.update_xaxes(showgrid=True, gridcolor=GRID)
                fig_gen_ref.update_yaxes(showgrid=False)
                st.plotly_chart(fig_gen_ref, use_container_width=True)
        else:
            st.info("Gender data not available in the Newly Registered sheet.")

        st.markdown("---")
        st.markdown("#### Dropped Volunteers — Reason Breakdown")
        st.caption("Source: Dropped VT sheet — volunteers who exited their offering(s).")
        if not dropped.empty and "Reasons" in dropped.columns:
            reason_counts = (
                dropped["Reasons"].fillna("Not specified")
                .value_counts().reset_index()
            )
            reason_counts.columns = ["Reason", "Count"]
            fig_drop = px.bar(
                reason_counts, x="Count", y="Reason",
                orientation="h", text="Count",
                color="Count",
                color_continuous_scale=[P["salmon"], P["red"]],
            )
            fig_drop.update_traces(
                textposition="outside", marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Count: %{x}<extra></extra>",
            )
            fig_drop.update_coloraxes(showscale=False)
            _layout(fig_drop, height=max(220, len(reason_counts) * 42),
                    margin=dict(l=0, r=60, t=20, b=0))
            fig_drop.update_layout(xaxis_title="Volunteers", yaxis_title="")
            fig_drop.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_drop.update_yaxes(showgrid=False)
            st.plotly_chart(fig_drop, use_container_width=True)
        else:
            st.info("No dropped volunteer records in the current data.")

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 2 — CENTRES
    # ═════════════════════════════════════════════════════════════════════════
    with tab_ctr:

        if df.empty:
            st.info(
                "🔍 No data available for the selected filters.\n\n"
                "Try broadening your Donor, State, or Subject selection in the sidebar."
            )
            st.stop()

        st.markdown("#### Centres & Volunteers by Donor")
        has_vrm_gender = {"En Boys", "En Girls"}.issubset(df.columns)
        agg_dict = dict(
            Centres    =("Center name",  "nunique"),
            Enrolled   =("Enrolled",     "sum"),
            CLH        =("CLH",          "sum"),
            Volunteers =("Volunteer id", "nunique"),
            Completed  =("Completed",    "sum"),
            Planned    =("Planned",      "sum"),
        )
        if has_vrm_gender:
            agg_dict["En_Boys"]  = ("En Boys",  "sum")
            agg_dict["En_Girls"] = ("En Girls", "sum")
        donor_summary = df.groupby("Donor").agg(**agg_dict).reset_index()
        donor_summary["Completion %"] = (
            donor_summary["Completed"]
            / donor_summary["Planned"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        donor_summary = donor_summary.sort_values("Centres", ascending=False)

        fig_donor = go.Figure()
        fig_donor.add_trace(go.Bar(
            name="Centres", x=donor_summary["Donor"], y=donor_summary["Centres"],
            marker_color=P["teal"], text=donor_summary["Centres"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Centres: %{y}<extra></extra>",
        ))
        fig_donor.add_trace(go.Bar(
            name="Volunteers", x=donor_summary["Donor"], y=donor_summary["Volunteers"],
            marker_color=P["green"], text=donor_summary["Volunteers"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Volunteers: %{y}<extra></extra>",
        ))
        fig_donor.update_layout(
            barmode="group", height=380,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=30, b=80),
            legend=dict(orientation="h", yanchor="top", y=-0.22,
                        xanchor="center", x=0.5, title_text=""),
            xaxis=dict(tickangle=-20, showgrid=False, linecolor="#dee2e6"),
            yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        st.plotly_chart(fig_donor, use_container_width=True)

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("##### Enrolled Students by Donor")
            fig_enr = px.bar(
                donor_summary.sort_values("Enrolled", ascending=True),
                x="Enrolled", y="Donor", orientation="h", text="Enrolled",
                color="Enrolled",
                color_continuous_scale=[P["lavender"], P["violet"]],
            )
            fig_enr.update_traces(textposition="outside",
                                  texttemplate="%{text:,}", marker_line_width=0)
            fig_enr.update_coloraxes(showscale=False)
            _layout(fig_enr, height=280, margin=dict(l=0, r=70, t=20, b=0))
            fig_enr.update_layout(xaxis_title="Enrolled Students", yaxis_title="")
            fig_enr.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_enr.update_yaxes(showgrid=False)
            st.plotly_chart(fig_enr, use_container_width=True)

        with col_d2:
            st.markdown("##### CLH by Donor")
            fig_clh_d = px.bar(
                donor_summary.sort_values("CLH", ascending=True),
                x="CLH", y="Donor", orientation="h", text="CLH",
                color="CLH",
                color_continuous_scale=[P["salmon"], P["coral"]],
            )
            fig_clh_d.update_traces(textposition="outside",
                                    texttemplate="%{text:,}", marker_line_width=0)
            fig_clh_d.update_coloraxes(showscale=False)
            _layout(fig_clh_d, height=280, margin=dict(l=0, r=70, t=20, b=0))
            fig_clh_d.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
            fig_clh_d.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_clh_d.update_yaxes(showgrid=False)
            st.plotly_chart(fig_clh_d, use_container_width=True)

        st.markdown("---")

        st.markdown("#### Enrolled Students — Gender Split by Donor")
        if has_vrm_gender:
            st.caption("En Boys and En Girls columns are available in this dataset.")
            gen_donor_melt = donor_summary[["Donor", "En_Boys", "En_Girls"]].melt(
                id_vars="Donor", value_vars=["En_Boys", "En_Girls"],
                var_name="Gender", value_name="Students"
            )
            gen_donor_melt["Gender"] = gen_donor_melt["Gender"].map(
                {"En_Boys": "Boys", "En_Girls": "Girls"}
            )
            fig_gen_enr = px.bar(
                gen_donor_melt, x="Donor", y="Students",
                color="Gender", barmode="group",
                text="Students",
                color_discrete_map={"Boys": P["teal"], "Girls": P["coral"]},
            )
            fig_gen_enr.update_traces(
                textposition="outside", texttemplate="%{text:,}",
                hovertemplate="<b>%{x}</b> · %{fullData.name}: %{y:,}<extra></extra>",
            )
            _layout(fig_gen_enr, height=340, legend_bottom=True,
                    margin=dict(l=0, r=0, t=20, b=60))
            fig_gen_enr.update_layout(xaxis_title="", yaxis_title="Enrolled Students")
            fig_gen_enr.update_xaxes(tickangle=-15, showgrid=False)
            st.plotly_chart(fig_gen_enr, use_container_width=True)
        else:
            st.info(
                "🔍 Gender breakdown (En Boys / En Girls) isn't available in this "
                f"month's ({selected_month}) VRM export, so this chart is skipped."
            )

        st.markdown("---")

        st.markdown("#### Top States by Active Centres")
        state_centres = (
            df.groupby("State").agg(
                Centres    =("Center name",  "nunique"),
                Enrolled   =("Enrolled",     "sum"),
                CLH        =("CLH",          "sum"),
                Volunteers =("Volunteer id", "nunique"),
            ).reset_index()
            .sort_values("Centres", ascending=True)
        )
        fig_sc = px.bar(
            state_centres, x="Centres", y="State",
            orientation="h", text="Centres",
            color="Centres",
            color_continuous_scale=[P["sky"], P["teal"]],
            hover_data={"Enrolled": ":,", "CLH": ":,", "Volunteers": True},
        )
        fig_sc.update_traces(textposition="outside", marker_line_width=0)
        fig_sc.update_coloraxes(showscale=False)
        _layout(fig_sc, height=max(260, len(state_centres) * 42),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_sc.update_layout(xaxis_title="Active Centres", yaxis_title="")
        fig_sc.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_sc.update_yaxes(showgrid=False)
        st.plotly_chart(fig_sc, use_container_width=True)

        st.markdown("---")

        st.markdown("#### Class Completion Split by Subject")
        st.caption("🟢 Completed  🟡 Offline (rescheduled / async)  🔴 Cancelled")
        exec_df = (
            df.groupby("Subject_clean")[["Completed", "Offline", "Cancelled"]]
            .sum().reset_index()
            .sort_values("Completed", ascending=False)
        )
        exec_melt = exec_df.melt(
            id_vars="Subject_clean",
            value_vars=["Completed", "Offline", "Cancelled"],
            var_name="Status", value_name="Sessions",
        )
        fig_exec = px.bar(
            exec_melt, x="Sessions", y="Subject_clean",
            color="Status", orientation="h", barmode="stack",
            text="Sessions",
            color_discrete_map={
                "Completed": P["green"],
                "Offline":   P["amber"],
                "Cancelled": P["red"],
            },
        )
        fig_exec.update_traces(
            textposition="inside", texttemplate="%{text:,}",
            hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
        )
        _layout(fig_exec, height=max(320, len(exec_df) * 42),
                legend_bottom=True, margin=dict(l=0, r=0, t=20, b=60))
        fig_exec.update_layout(xaxis_title="Sessions", yaxis_title="")
        fig_exec.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_exec.update_yaxes(showgrid=False)
        st.plotly_chart(fig_exec, use_container_width=True)

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 3 — ACADEMIC HEALTH
    # ═════════════════════════════════════════════════════════════════════════
    with tab_aca:

        if df.empty:
            st.info(
                "🔍 No data available for the selected filters.\n\n"
                "Try broadening your Donor, State, or Subject selection in the sidebar."
            )
            st.stop()

        st.markdown("#### CLH by Subject")
        clh_subj = (
            df.groupby("Subject_clean")["CLH"].sum()
            .reset_index()
            .sort_values("CLH", ascending=True)
        )
        fig_clh = px.bar(
            clh_subj, x="CLH", y="Subject_clean",
            orientation="h", text="CLH",
            color="CLH",
            color_continuous_scale=[P["salmon"], P["coral"]],
        )
        fig_clh.update_traces(
            textposition="outside", texttemplate="%{text:,}",
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>CLH: %{x:,}<extra></extra>",
        )
        fig_clh.update_coloraxes(showscale=False)
        _layout(fig_clh, height=max(300, len(clh_subj) * 40),
                margin=dict(l=0, r=70, t=20, b=0))
        fig_clh.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
        fig_clh.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_clh.update_yaxes(showgrid=False)
        st.plotly_chart(fig_clh, use_container_width=True)

        st.markdown("---")

        st.markdown("#### Average Attendance % — State × Subject Heatmap")
        st.caption("Scale clamped 50–100%. White = no sessions for that pairing.")
        heat_df = (
            df.groupby(["State", "Subject_clean"])["Attendance%"]
            .mean().reset_index()
        )
        heat_pivot = (
            heat_df.pivot(index="State", columns="Subject_clean", values="Attendance%")
            .fillna(0)
        )
        sort_order = (
            heat_pivot.replace(0, pd.NA).mean(axis=1)
            .sort_values(ascending=False).index
        )
        heat_pivot = heat_pivot.loc[sort_order]

        fig_heat = px.imshow(
            heat_pivot,
            color_continuous_scale=["#dfe6e9", P["teal"]],
            aspect="auto",
            text_auto=".0f",
            zmin=50, zmax=100,
        )
        fig_heat.update_traces(
            hovertemplate=(
                "<b>%{y}</b> · <b>%{x}</b><br>"
                "Avg Attendance: %{z:.1f}%<extra></extra>"
            )
        )
        fig_heat.update_layout(
            height=max(300, len(heat_pivot) * 50),
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=20, b=120),
            coloraxis_colorbar=dict(
                title="Att%", ticksuffix="%", len=0.6, thickness=12,
            ),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_heat.update_xaxes(tickangle=-40, side="bottom", showgrid=False)
        fig_heat.update_yaxes(showgrid=False)
        st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown("---")

        st.markdown("#### Attendance Distribution Across All Sessions")
        fig_hist = px.histogram(
            df, x="Attendance%", nbins=25,
            color_discrete_sequence=[P["violet"]], opacity=0.82,
        )
        fig_hist.update_layout(
            xaxis_title="Attendance %", yaxis_title="Number of Sessions",
            height=240,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=20, b=0),
            bargap=0.05,
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_hist.update_xaxes(showgrid=False)
        fig_hist.update_yaxes(showgrid=True, gridcolor=GRID)
        st.plotly_chart(fig_hist, use_container_width=True)

        st.markdown("---")

        st.markdown("#### 🗂️ Filtered Operational Data")
        st.caption(
            f"**{len(df):,} rows** match current filters · sorted by CLH · "
            "click any column header to re-sort."
        )
        display_cols = [
            "Volunteer name", "Center name", "State", "Donor",
            "Subject_clean", "Registered", "Reg Boys", "Reg Girls",
            "Enrolled", "En Boys", "En Girls",
            "Attendance%", "CLH", "Total hours(Comp+Offline)",
            "Planned", "Completed", "Offline", "Cancelled",
        ]
        available_display = [c for c in display_cols if c in df.columns]
        display_df = df[available_display].rename(columns={
            "Subject_clean":             "Subject",
            "Total hours(Comp+Offline)": "Vol Hrs",
        }).copy()
        display_df["Attendance%"] = display_df["Attendance%"].round(1)
        st.dataframe(
            display_df.sort_values("CLH", ascending=False),
            use_container_width=True,
            hide_index=True,
            height=420,
        )

    # ═════════════════════════════════════════════════════════════════════════

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 4 — DRM CLIENT REPORT  (premium client-facing tab)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:

        if not drm_data_raw:
            st.error(
                f"⚠️ No DRM file found for **{selected_month}**. "
                f"Place a file named like `DRM_{selected_month.replace(' ', '_')}.xlsx` "
                f"in `{DATA_DIR}` to enable this report."
            )
            st.stop()

        sess_raw = drm_data_raw["sess"]
        drm_raw  = drm_data_raw["drm"]
        od_raw   = drm_data_raw["od"]

        # Apply DRM filters (sel_drm_* come from the sidebar built above)
        sess = sess_raw[
            sess_raw["Donor"].isin(sel_drm_donors) &
            sess_raw["State"].isin(sel_drm_states) &
            sess_raw["Subject_clean"].isin(sel_drm_subjects)
        ].copy()

        drm = drm_raw[
            drm_raw["Donor Name"].isin(sel_drm_donors) &
            drm_raw["State"].isin(sel_drm_states)
        ].copy()

        od = od_raw[
            od_raw["Donor"].isin(sel_drm_donors) &
            od_raw["State"].isin(sel_drm_states) &
            od_raw["Subject_clean"].isin(sel_drm_subjects)
        ].copy()

        # ── Guard: empty result after filtering ──────────────────────────────────
        if sess.empty or drm.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.info(
                "🔍 **No data available for the selected filters.**\n\n"
                "The combination of Donor, State, and Subject you selected returned no sessions "
                "or no active centres. Try:\n"
                "- Selecting additional donors or states\n"
                "- Switching to **DRM Filters** in the sidebar and broadening your selection"
            )
            st.stop()

        # ── Report Header ──────────────────────────────────────────────────────
        num_states_live   = drm["State"].nunique()
        num_centres_live  = drm["Center Name"].nunique()
        num_donors_live   = drm["Donor Name"].nunique()
        st.markdown(
            f"""
            <div style='background:linear-gradient(135deg,#0f8a6e 0%,#185fa5 100%);
                        border-radius:14px;padding:26px 32px;margin-bottom:24px;'>
                <div style='display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;'>
                    <div>
                        <h2 style='color:white;margin:0;font-size:1.7rem;font-weight:700;letter-spacing:-0.02em;'>
                            📊 eVidyaloka — Programme Impact Report
                        </h2>
                        <p style='color:rgba(255,255,255,0.82);margin:6px 0 0 0;font-size:0.95em;'>
                            Donor Relationship Management &nbsp;·&nbsp; Session Analytics
                            &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {selected_month}
                        </p>
                    </div>
                    <div style='display:flex;gap:8px;flex-wrap:wrap;'>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>{num_centres_live} Centres</span>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>{num_states_live} States</span>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>{num_donors_live} Donors</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 1 — HEADLINE KPIs  (linked to DRM filters)
        # ═══════════════════════════════════════════════════════════════════
        total_sessions = len(sess)
        comp_sessions  = int((sess["Session_status"] == "Completed").sum())
        off_sessions   = int((sess["Session_status"] == "Offline").sum())
        canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
        total_clh_drm  = int(sess["Present/CLH"].sum())
        total_students = int(drm["Enrolled"].sum())
        avg_att_drm    = sess["Attendance%"].mean()
        total_centres  = drm["Center Name"].nunique()
        drm_total_vols = sess["Volunteer_id"].nunique()
        comp_rate      = (comp_sessions + off_sessions) / total_sessions * 100 if total_sessions else 0
        canc_rate      = canc_sessions / total_sessions * 100 if total_sessions else 0
        total_girls    = int(drm["En Girls"].sum())
        total_boys     = int(drm["En Boys"].sum())
        total_enrolled_all = total_boys + total_girls
        girl_pct       = total_girls / total_enrolled_all * 100 if total_enrolled_all else 0
        boy_pct        = 100 - girl_pct
        num_states     = drm["State"].nunique()

        st.markdown("#### 🎯 Programme Snapshot")

        r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
        r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns(5)

        # Row 1 — programme reach
        for col_w, lbl, val, color, sub in [
            (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
            (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique this month"),
            (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned this month"),
            (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
            (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Row 2 — impact metrics
        # "Enrolled Students" card now shows Boys / Girls % inline as a mini split
        boy_bar    = int(boy_pct)
        girl_bar   = int(girl_pct)
        _c_teal    = D["teal"]
        _c_blue    = D["blue"]
        _c_coral   = D["coral"]
        gender_card = (
            f"<div style='background:white;border-radius:12px;padding:18px 14px;"
            f"border-left:5px solid {_c_teal};text-align:center;"
            f"box-shadow:0 2px 8px rgba(0,0,0,0.06);'>"
            f"<div style='font-size:0.68em;color:#636e72;font-weight:700;"
            f"letter-spacing:0.07em;text-transform:uppercase;margin-bottom:5px;'>Enrolled Students</div>"
            f"<div style='font-size:1.9rem;font-weight:800;color:{_c_teal};"
            f"line-height:1.1;font-family:DM Mono,monospace;'>{total_enrolled_all:,}</div>"
            f"<div style='display:flex;height:6px;border-radius:3px;overflow:hidden;margin:8px 0 5px;'>"
            f"<div style='width:{boy_bar}%;background:{_c_blue};'></div>"
            f"<div style='width:{girl_bar}%;background:{_c_coral};'></div>"
            f"</div>"
            f"<div style='display:flex;justify-content:center;gap:14px;font-size:0.68em;color:#636e72;'>"
            f"<span>&#128102; Boys {boy_pct:.0f}%</span><span>&#128103; Girls {girl_pct:.0f}%</span>"
            f"</div></div>"
        )
        r2c2.markdown(gender_card, unsafe_allow_html=True)

        for col_w, lbl, val, color, sub in [
            (r2c1, "Total CLH",       f"{total_clh_drm:,}",   D["purple"], "child learning hours"),
            (r2c3, "Avg Attendance",  f"{avg_att_drm:.1f}%",  D["amber"],  "session attendance"),
            (r2c4, "Completion Rate", f"{comp_rate:.1f}%",    D["green"],  "completed + offline"),
            (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{off_sessions/total_sessions*100:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 2 — STATE-WISE CENTRES + ENROLLED STUDENTS
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### 🗺️ State-wise Centres & Enrolled Students")

        state_enr = (
            drm.groupby("State").agg(
                Centres  =("Center Name",    "count"),
                Enrolled =("Enrolled",       "sum"),
                Boys     =("En Boys",        "sum"),
                Girls    =("En Girls",       "sum"),
                CLH      =("Live_CLH",       "sum"),
                Vols     =("Live Volunteers","sum"),
            ).reset_index()
        )
        state_enr["Girl%"] = (
            state_enr["Girls"] / (state_enr["Boys"] + state_enr["Girls"]).replace(0, pd.NA) * 100
        ).round(1).fillna(0)
        state_enr = state_enr.sort_values("Enrolled", ascending=True)

        col_se1, col_se2 = st.columns(2)

        with col_se1:
            st.markdown("##### Centres per state")
            fig_sc = px.bar(
                state_enr, x="Centres", y="State",
                orientation="h", text="Centres",
                color="Centres",
                color_continuous_scale=[[0, "#cce0f5"], [1, D["blue"]]],
                hover_data={"Enrolled": ":,", "Vols": True},
            )
            fig_sc.update_traces(textposition="outside", marker_line_width=0,
                                 hovertemplate="<b>%{y}</b><br>Centres: %{x}<br>Enrolled: %{customdata[0]:,}<extra></extra>")
            fig_sc.update_coloraxes(showscale=False)
            _layout(fig_sc, height=280, margin=dict(l=0, r=50, t=10, b=0))
            fig_sc.update_layout(xaxis_title="Centres", yaxis_title="")
            fig_sc.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_sc.update_yaxes(showgrid=False)
            st.plotly_chart(fig_sc, use_container_width=True)

        with col_se2:
            st.markdown("##### Enrolled students per state — boys vs girls")
            state_gen_melt = state_enr.melt(
                id_vars="State", value_vars=["Boys","Girls"],
                var_name="Gender", value_name="Students"
            )
            fig_sgen = px.bar(
                state_gen_melt, x="Students", y="State",
                color="Gender", barmode="stack", orientation="h",
                text="Students",
                color_discrete_map={"Boys": D["blue"], "Girls": D["coral"]},
            )
            fig_sgen.update_traces(textposition="inside", texttemplate="%{text:,}",
                                   hovertemplate="<b>%{y}</b> · %{fullData.name}: %{x:,}<extra></extra>")
            _layout(fig_sgen, height=280, legend_bottom=True, margin=dict(l=0, r=0, t=10, b=50))
            fig_sgen.update_layout(xaxis_title="Enrolled Students", yaxis_title="")
            fig_sgen.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_sgen.update_yaxes(showgrid=False)
            st.plotly_chart(fig_sgen, use_container_width=True)

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 3 — STATE-WISE SESSION STATUS + ATTENDANCE
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### 📋 State-wise Session Status & Attendance")

        state_sess = (
            sess.groupby("State").agg(
                Completed =("Session_status", lambda x: (x == "Completed").sum()),
                Offline   =("Session_status", lambda x: (x == "Offline").sum()),
                Cancelled =("Session_status", lambda x: (x == "Cancelled").sum()),
                Attendance=("Attendance%",    "mean"),
                CLH       =("Present/CLH",    "sum"),
            ).reset_index()
        )
        state_sess["Total"] = state_sess["Completed"] + state_sess["Offline"] + state_sess["Cancelled"]
        state_sess["Comp%"] = (
            (state_sess["Completed"] + state_sess["Offline"])
            / state_sess["Total"].replace(0, pd.NA) * 100
        ).round(1).fillna(0)
        state_sess = state_sess.sort_values("Total", ascending=True)

        col_ss1, col_ss2 = st.columns([1.5, 1])

        with col_ss1:
            st.markdown("##### Session status by state")
            state_melt = state_sess.melt(
                id_vars=["State", "Attendance"],
                value_vars=["Completed", "Offline", "Cancelled"],
                var_name="Status", value_name="Sessions"
            )
            fig_sstat = px.bar(
                state_melt, x="Sessions", y="State",
                color="Status", barmode="stack", orientation="h",
                text="Sessions",
                color_discrete_map={
                    "Completed": D["teal"],
                    "Offline":   D["amber"],
                    "Cancelled": D["red"],
                },
                category_orders={"State": state_sess["State"].tolist()},
            )
            fig_sstat.update_traces(textposition="inside", texttemplate="%{text}",
                                    hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x}<extra></extra>")
            _layout(fig_sstat, height=320, legend_bottom=True, margin=dict(l=0, r=0, t=10, b=50))
            fig_sstat.update_layout(xaxis_title="Sessions", yaxis_title="")
            fig_sstat.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_sstat.update_yaxes(showgrid=False)
            st.plotly_chart(fig_sstat, use_container_width=True)

        with col_ss2:
            st.markdown("##### Avg attendance % by state")
            fig_satt = px.bar(
                state_sess.sort_values("Attendance", ascending=True),
                x="Attendance", y="State",
                orientation="h",
                text=state_sess.sort_values("Attendance", ascending=True)["Attendance"].round(1).astype(str) + "%",
                color="Attendance",
                color_continuous_scale=["#e17055", D["amber"], D["teal"]],
                range_color=[60, 100],
            )
            fig_satt.update_traces(textposition="outside", marker_line_width=0,
                                   hovertemplate="<b>%{y}</b><br>Attendance: %{x:.1f}%<extra></extra>")
            fig_satt.update_coloraxes(showscale=False)
            _layout(fig_satt, height=320, margin=dict(l=0, r=70, t=10, b=0))
            fig_satt.update_layout(xaxis_title="Avg Attendance %", yaxis_title="",
                                   xaxis=dict(range=[50, 115], ticksuffix="%", showgrid=True, gridcolor=GRID))
            fig_satt.update_yaxes(showgrid=False)
            st.plotly_chart(fig_satt, use_container_width=True)

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 4 — CANCELLATION REASONS + STATE-WISE CLH SPLIT
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### ❌ Cancellations & CLH Split")

        col_ca1, col_ca2 = st.columns([1, 1.4])

        with col_ca1:
            st.markdown("##### Cancellation reasons")
            canc_df = (
                sess[sess["Session_status"] == "Cancelled"]
                ["Cancel_reason"].fillna("Not specified")
                .value_counts().reset_index()
            )
            canc_df.columns = ["Reason", "Count"]
            total_canc = canc_df["Count"].sum()

            if not canc_df.empty:
                fig_canc = px.pie(
                    canc_df, names="Reason", values="Count",
                    hole=0.52,
                    color_discrete_sequence=[D["coral"], D["amber"], D["red"], D["purple"]],
                )
                fig_canc.update_traces(
                    texttemplate="<b>%{label}</b><br>%{value}  (%{percent:.0%})",
                    textposition="outside",
                    hovertemplate="<b>%{label}</b><br>Sessions: %{value}<extra></extra>",
                    pull=[0.04] * len(canc_df),
                )
                fig_canc.update_layout(
                    height=300, showlegend=False,
                    plot_bgcolor=BG, paper_bgcolor=BG,
                    margin=dict(l=10, r=10, t=20, b=10),
                    annotations=[dict(
                        text=f"<b>{total_canc}</b><br>cancelled",
                        x=0.5, y=0.5, font_size=14, showarrow=False,
                        font=dict(color="#2d3436"),
                    )],
                )
                st.plotly_chart(fig_canc, use_container_width=True)

                top_reason = canc_df.iloc[0]["Reason"]
                top_pct    = canc_df.iloc[0]["Count"] / total_canc * 100
                st.markdown(
                    _insight_box(
                        f"<b>{top_pct:.0f}%</b> of cancellations are due to <b>{top_reason}</b>. "
                        "Proactive scheduling support could recover these sessions."
                    ),
                    unsafe_allow_html=True,
                )
            else:
                st.info("No cancellations in the current filter selection.")

        with col_ca2:
            st.markdown("##### State-wise CLH split")
            # Stacked bar: each state bar is split by donor CLH contribution
            state_donor_clh = (
                sess.groupby(["State", "Donor"])["Present/CLH"]
                .sum().reset_index()
                .rename(columns={"Present/CLH": "CLH"})
            )
            state_totals = state_donor_clh.groupby("State")["CLH"].sum().sort_values(ascending=True)
            state_order  = state_totals.index.tolist()

            fig_sclh = px.bar(
                state_donor_clh, x="CLH", y="State",
                color="Donor", barmode="stack", orientation="h",
                text="CLH",
                color_discrete_sequence=[D["blue"], D["teal"], D["purple"], D["amber"], D["gray"]],
                category_orders={"State": state_order},
            )
            fig_sclh.update_traces(
                texttemplate="%{text:,}", textposition="inside",
                hovertemplate="<b>%{y}</b> · %{fullData.name}<br>CLH: %{x:,}<extra></extra>",
            )
            _layout(fig_sclh, height=300, legend_bottom=True, margin=dict(l=0, r=0, t=10, b=50))
            fig_sclh.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
            fig_sclh.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_sclh.update_yaxes(showgrid=False)
            st.plotly_chart(fig_sclh, use_container_width=True)

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 5 — SESSIONS BY SUBJECT & GRADE
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### 📚 Subject & Grade Analytics")

        # --- 5a: Sessions heatmap subject × grade ---
        col_sg1, col_sg2 = st.columns(2)

        with col_sg1:
            st.markdown("##### Sessions by subject & grade")
            subj_grade_sess = (
                sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                .count().reset_index()
                .rename(columns={"Session_id": "Sessions"})
            )
            pivot_sgs = (
                subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                .fillna(0)
            )
            # Sort rows by total sessions descending
            pivot_sgs = pivot_sgs.loc[pivot_sgs.sum(axis=1).sort_values(ascending=False).index]
            fig_sgs = px.imshow(
                pivot_sgs,
                color_continuous_scale=["#e1f5ee", D["teal"]],
                aspect="auto", text_auto=".0f",
            )
            fig_sgs.update_traces(
                hovertemplate="<b>%{y}</b> · Grade %{x}<br>Sessions: %{z:.0f}<extra></extra>"
            )
            fig_sgs.update_layout(
                height=300, plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=0, t=20, b=0),
                coloraxis_showscale=False,
                xaxis=dict(title="Grade", showgrid=False),
                yaxis=dict(title="", showgrid=False),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
            )
            st.plotly_chart(fig_sgs, use_container_width=True)

        # --- 5b: Subject × Grade session completion heatmap (% completed) ---
        with col_sg2:
            st.markdown("##### Session completion % — subject & grade")
            subj_grade_comp = (
                sess.groupby(["Subject_clean", "Grade"]).agg(
                    Total    =("Session_id",     "count"),
                    Done     =("Session_status", lambda x: ((x == "Completed") | (x == "Offline")).sum()),
                ).reset_index()
            )
            subj_grade_comp["Comp%"] = (
                subj_grade_comp["Done"] / subj_grade_comp["Total"].replace(0, pd.NA) * 100
            ).round(1).fillna(0)
            pivot_comp = (
                subj_grade_comp.pivot(index="Subject_clean", columns="Grade", values="Comp%")
                .fillna(0)
            )
            # Match row order of session heatmap
            common_idx = [i for i in pivot_sgs.index if i in pivot_comp.index]
            pivot_comp = pivot_comp.reindex(common_idx)
            fig_comp = px.imshow(
                pivot_comp,
                color_continuous_scale=["#faeeda", D["teal"]],
                aspect="auto", text_auto=".0f",
                zmin=0, zmax=100,
            )
            fig_comp.update_traces(
                hovertemplate="<b>%{y}</b> · Grade %{x}<br>Completion: %{z:.1f}%<extra></extra>"
            )
            fig_comp.update_layout(
                height=300, plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=0, t=20, b=0),
                coloraxis_showscale=False,
                xaxis=dict(title="Grade", showgrid=False),
                yaxis=dict(title="", showgrid=False),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
            )
            st.plotly_chart(fig_comp, use_container_width=True)

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 6 — CLH BY SUBJECT & GRADE
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### ⏱️ CLH by Subject & Grade")

        col_clh1, col_clh2 = st.columns([1.4, 1])

        with col_clh1:
            st.markdown("##### CLH heatmap — subject × grade")
            subj_grade_clh = (
                sess.groupby(["Subject_clean", "Grade"])["Present/CLH"]
                .sum().reset_index()
                .rename(columns={"Present/CLH": "CLH"})
            )
            pivot_clhsg = (
                subj_grade_clh.pivot(index="Subject_clean", columns="Grade", values="CLH")
                .fillna(0)
            )
            pivot_clhsg = pivot_clhsg.loc[pivot_clhsg.sum(axis=1).sort_values(ascending=False).index]
            fig_clhsg = px.imshow(
                pivot_clhsg,
                color_continuous_scale=["#e6f1fb", D["blue"]],
                aspect="auto", text_auto=".0f",
            )
            fig_clhsg.update_traces(
                hovertemplate="<b>%{y}</b> · Grade %{x}<br>CLH: %{z:,.0f}<extra></extra>"
            )
            fig_clhsg.update_layout(
                height=300, plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=0, t=20, b=0),
                coloraxis_showscale=False,
                xaxis=dict(title="Grade", showgrid=False),
                yaxis=dict(title="", showgrid=False),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
            )
            st.plotly_chart(fig_clhsg, use_container_width=True)

        with col_clh2:
            st.markdown("##### CLH by subject — total")
            clh_by_subj = (
                sess.groupby("Subject_clean")["Present/CLH"]
                .sum().reset_index()
                .rename(columns={"Present/CLH": "CLH"})
                .sort_values("CLH", ascending=True)
            )
            fig_clhs = px.bar(
                clh_by_subj, x="CLH", y="Subject_clean",
                orientation="h", text="CLH",
                color="CLH",
                color_continuous_scale=[[0, "#cce0f5"], [1, D["blue"]]],
            )
            fig_clhs.update_traces(
                texttemplate="%{text:,}", textposition="outside",
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>CLH: %{x:,}<extra></extra>",
            )
            fig_clhs.update_coloraxes(showscale=False)
            _layout(fig_clhs, height=300, margin=dict(l=0, r=70, t=20, b=0))
            fig_clhs.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
            fig_clhs.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_clhs.update_yaxes(showgrid=False)
            st.plotly_chart(fig_clhs, use_container_width=True)

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 7 — VOLUNTEER PERFORMANCE SCORECARD
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### 🙋 Volunteer Performance Scorecard")

        vol_perf = (
            sess.groupby(["Volunteer_id", "Teacher_name"]).agg(
                Sessions  =("Session_id",     "count"),
                Completed =("Session_status", lambda x: (x == "Completed").sum()),
                CLH       =("Present/CLH",    "sum"),
                Avg_Att   =("Attendance%",    "mean"),
                Students  =("Total students", "sum"),
                Subjects  =("Subject_clean",  lambda x: ", ".join(sorted(x.unique()))),
            ).reset_index()
        )
        vol_perf["Comp %"]  = (vol_perf["Completed"] / vol_perf["Sessions"] * 100).round(1)
        vol_perf["Avg_Att"] = vol_perf["Avg_Att"].round(1)

        def _flag(row):
            if row["Comp %"] < 60:  return "🔴 At Risk"
            if row["Comp %"] < 85:  return "🟡 Moderate"
            return "🟢 Strong"

        vol_perf["Status"] = vol_perf.apply(_flag, axis=1)
        vol_display = vol_perf[[
            "Teacher_name", "Subjects", "Sessions", "Comp %",
            "Avg_Att", "CLH", "Students", "Status"
        ]].rename(columns={
            "Teacher_name": "Volunteer",
            "Avg_Att":      "Att %",
        }).sort_values("CLH", ascending=False)

        st.dataframe(
            vol_display,
            use_container_width=True,
            hide_index=True,
            height=400,
            column_config={
                "CLH":    st.column_config.NumberColumn("CLH",    format="%d"),
                "Comp %": st.column_config.ProgressColumn(
                    "Comp %", min_value=0, max_value=100, format="%.1f%%"
                ),
                "Att %":  st.column_config.ProgressColumn(
                    "Att %",  min_value=0, max_value=100, format="%.1f%%"
                ),
            },
        )

        st.markdown("---")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION 8 — FULL CENTRE SCORECARD TABLE
        # ═══════════════════════════════════════════════════════════════════
        st.markdown("#### 🗂️ Full Centre Scorecard")
        st.caption(
            f"All {total_centres} active centres ranked by CLH delivered. &nbsp;"
            "🟢 ≥ 90% completion &nbsp; 🟡 70–90% &nbsp; 🔴 < 70% &nbsp;|&nbsp; "
            "Click any column header to re-sort."
        )
        drm_display = drm[[
            "Center Name", "State", "Donor Name",
            "Live Volunteers", "Planned", "Completed", "Offline", "Cancelled",
            "Completion%", "Cancellation%",
            "Live_CLH", "Enrolled", "En Boys", "En Girls",
            "Attendance %", "Girl%"
        ]].copy().rename(columns={
            "Donor Name":     "Donor",
            "Live Volunteers":"Vols",
            "Completion%":    "Comp %",
            "Cancellation%":  "Canc %",
            "Live_CLH":       "CLH",
            "Attendance %":   "Att %",
            "Girl%":          "Girl %",
        })
        drm_display["Att %"]  = drm_display["Att %"].round(1)
        drm_display["Comp %"] = drm_display["Comp %"].round(1)
        drm_display["Girl %"] = drm_display["Girl %"].round(1)
        drm_display["Canc %"] = drm_display["Canc %"].round(1)
        drm_display = drm_display.sort_values("CLH", ascending=False)

        st.dataframe(
            drm_display,
            use_container_width=True,
            hide_index=True,
            height=520,
            column_config={
                "Comp %": st.column_config.ProgressColumn(
                    "Comp %", min_value=0, max_value=100, format="%.1f%%"
                ),
                "Att %":  st.column_config.ProgressColumn(
                    "Att %",  min_value=0, max_value=100, format="%.1f%%"
                ),
                "CLH":    st.column_config.NumberColumn("CLH",    format="%d"),
                "Girl %": st.column_config.NumberColumn("Girl %", format="%.1f%%"),
                "Canc %": st.column_config.NumberColumn("Canc %", format="%.1f%%"),
            },
        )

        # ── Footer ─────────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                    padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                eVidyaloka Programme Impact Report &nbsp;·&nbsp; {selected_month}
                &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                {drm_total_vols} volunteers &nbsp;·&nbsp;
                {total_students:,} students &nbsp;·&nbsp;
                {total_clh_drm:,} child learning hours
            </div>""",
            unsafe_allow_html=True,
        )
    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — Topic & SubTopic Analytics
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("""
        <style>
        .metric-card{
            background:#ffffff; padding:18px; border-radius:12px;
            box-shadow:0 2px 10px rgba(0,0,0,.08);
        }
        [data-testid="stMetricValue"]{
            font-size:34px; font-weight:700; color:#0f766e;
        }
        </style>
        """, unsafe_allow_html=True)

        topic_df, cancelled_df, offline_df = load_topic_data()

        if topic_df is None:
            st.error(
                f"Topic/SubTopic data file not found. "
                f"Place `Topic_SubTopic_Cancelled_Offline_May2026.xlsx` in `{DATA_DIR}`."
            )
        else:
            st.title("📚 Topic & SubTopic Analytics Dashboard")

            # Filters in-tab (sidebar is used by VRM/DRM)
            with st.expander("📊 Filters", expanded=True):
                fc1, fc2, fc3 = st.columns(3)
                base_df = topic_df.copy()

                with fc1:
                    state_filter = st.multiselect(
                        "State",
                        sorted(base_df["state"].dropna().unique()),
                        key="top_state",
                    )
                state_df = base_df[base_df["state"].isin(state_filter)] if state_filter else base_df.copy()

                with fc2:
                    subject_filter = st.multiselect(
                        "Subject",
                        sorted(state_df["subject"].dropna().unique()),
                        key="top_subject",
                    )
                subject_df = state_df[state_df["subject"].isin(subject_filter)] if subject_filter else state_df.copy()

                with fc3:
                    grade_filter = st.multiselect(
                        "Grade",
                        sorted(subject_df["grade"].dropna().unique()),
                        key="top_grade",
                    )
                grade_df = subject_df[subject_df["grade"].isin(grade_filter)] if grade_filter else subject_df.copy()

                fc4, fc5 = st.columns(2)
                with fc4:
                    status_filter = st.multiselect(
                        "Session Status",
                        sorted(grade_df["session_status"].dropna().unique()),
                        key="top_status",
                    )
                status_df = grade_df[grade_df["session_status"].isin(status_filter)] if status_filter else grade_df.copy()

                with fc5:
                    centre_filter = st.multiselect(
                        "Centre",
                        sorted(status_df["center_name"].dropna().unique()),
                        key="top_centre",
                    )
                filtered_df = (
                    status_df[status_df["center_name"].isin(centre_filter)]
                    if centre_filter else status_df.copy()
                )

            # KPI cards
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Sessions", f"{len(filtered_df):,}")
            with col2:
                st.metric("Unique Topics", f"{filtered_df['topic_name'].nunique():,}")
            with col3:
                st.metric("Unique Subtopics", f"{filtered_df['sub_topic_name'].nunique():,}")
            with col4:
                st.metric("Active Centers", f"{filtered_df['center_id'].nunique():,}")

            st.subheader("📋 Topic / Sub-topic Session Summary")
            pivot_table = (
                filtered_df
                .groupby(["topic_name", "sub_topic_name"], dropna=False)
                .agg(**{
                    "#Sessions": ("topic_name", "count"),
                    "#Schools": ("center_name", "nunique"),
                })
                .reset_index()
                .sort_values("#Sessions", ascending=False)
                .rename(columns={"topic_name": "Topic", "sub_topic_name": "Sub-topic"})
            )
            st.dataframe(pivot_table, use_container_width=True, hide_index=True, height=700)

            st.download_button(
                "📥 Download Topic Data",
                filtered_df.to_csv(index=False),
                file_name="topic_data.csv",
                key="top_download",
            )
