
Gemini

New chat
Library
Images
Scheduled
Students
Customise
New notebook
Untitled notebook
Multi-Month Data Combination Functions
Monthly to Quarterly Data Reporting
Profession Mapping Table Generation
Download Power BI Desktop Guide
Conversation with Gemini
in earlier code you removed topic tab add it again

Here is the updated, complete ops_dashboard.py code with Tab 5 ("📚 Topic & SubTopic Analytics") added back alongside the Half-Yearly reporting feature.

Key Highlights Included:
Half-Yearly Option: Renders "Half-Yearly" in the top reporting radio button along with "Monthly" and "Quarterly".

Restored Tab 5: Re-integrated Tab 5 ("📚 Topic & SubTopic Analytics") with its file uploader, state/subject/grade filters, analytics views (Top Topics, Top SubTopics, Cancelled, Offline), and CSV export.

Python
"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,
Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().

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
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

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
    them by reporting month.
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


# ─────────────────────────────────────────────────────────────────────────────
# QUARTERLY & HALF-YEARLY REPORTING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_quarter_label(month_num: int, year: int) -> str:
    """Financial year quarters:
       Q1 = Apr-Jun
       Q2 = Jul-Sep
       Q3 = Oct-Dec
       Q4 = Jan-Mar
    """
    if month_num in (4, 5, 6):
        return f"Q1 {year} (Apr–Jun)"
    elif month_num in (7, 8, 9):
        return f"Q2 {year} (Jul–Sep)"
    elif month_num in (10, 11, 12):
        return f"Q3 {year} (Oct–Dec)"
    else:
        return f"Q4 {year - 1}-{year} (Jan–Mar)"


def discover_quarterly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial quarters."""
    monthly = discover_monthly_files(data_dir)
    quarters = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        quarter = get_quarter_label(month_num, year)
        if quarter not in quarters:
            quarters[quarter] = []

        quarters[quarter].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for quarter in quarters:
        quarters[quarter] = sorted(
            quarters[quarter],
            key=lambda x: (x["year"], x["month"])
        )

    return quarters


def get_half_yearly_label(month_num: int, year: int) -> str:
    """Financial year half-years:
       H1 = Apr-Sep
       H2 = Oct-Mar
    """
    if month_num in (4, 5, 6, 7, 8, 9):
        return f"H1 {year} (Apr–Sep)"
    else:
        fy_year = year if month_num in (10, 11, 12) else year - 1
        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"


def discover_half_yearly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial half-years."""
    monthly = discover_monthly_files(data_dir)
    half_years = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        hy_label = get_half_yearly_label(month_num, year)
        if hy_label not in half_years:
            half_years[hy_label] = []

        half_years[hy_label].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for hy in half_years:
        half_years[hy] = sorted(
            half_years[hy],
            key=lambda x: (x["year"], x["month"])
        )

    return half_years


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-MONTH VRM AND DRM COMBINE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_selected_vrm_data(selected_entries: list) -> dict:
    """
    Loads VRM data for multiple selected month entries and concatenates 
    them together across 'active', 'dropped', and 'new_reg' sheets.
    """
    active_dfs, dropped_dfs, new_reg_dfs = [], [], []

    for entry in selected_entries:
        vrm_path = entry.get("vrm")
        if vrm_path and os.path.exists(vrm_path):
            month_data = load_data(vrm_path)
            if month_data:
                if "active" in month_data and not month_data["active"].empty:
                    df_act = month_data["active"].copy()
                    df_act["Reporting Period"] = entry["label"]
                    active_dfs.append(df_act)
                
                if "dropped" in month_data and not month_data["dropped"].empty:
                    df_drp = month_data["dropped"].copy()
                    df_drp["Reporting Period"] = entry["label"]
                    dropped_dfs.append(df_drp)
                
                if "new_reg" in month_data and not month_data["new_reg"].empty:
                    df_nr = month_data["new_reg"].copy()
                    df_nr["Reporting Period"] = entry["label"]
                    new_reg_dfs.append(df_nr)

    return {
        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),
        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),
        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),
    }


def load_selected_drm_data(selected_entries: list) -> dict:
    """
    Loads DRM data for multiple selected month entries and concatenates 
    them across all individual DRM sheets.
    """
    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []

    for entry in selected_entries:
        drm_path = entry.get("drm")
        if drm_path and os.path.exists(drm_path):
            m_drm = load_drm_data(drm_path)
            if m_drm:
                for key, container in [
                    ("sess", sess_dfs),
                    ("drm", drm_dfs),
                    ("ac", ac_dfs),
                    ("od", od_dfs),
                    ("ne", ne_dfs),
                ]:
                    if key in m_drm and not m_drm[key].empty:
                        df_sheet = m_drm[key].copy()
                        df_sheet["Reporting Period"] = entry["label"]
                        container.append(df_sheet)

    return {
        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),
        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),
        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),
        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),
        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),
    }


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

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

    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))
    dropped.columns = dropped.columns.str.strip()
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

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

DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

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

    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))
    ac.columns = ac.columns.str.strip()

    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

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

    # ─────────────────────────────────────────────────────────────────────────
    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)
    # ─────────────────────────────────────────────────────────────────────────

    reporting_type = st.radio(
        "📊 Reporting Type",
        ["Monthly", "Quarterly", "Half-Yearly"],
        horizontal=True,
        key="reporting_type",
    )

    selected_entries = []
    period_label = ""

    if reporting_type == "Monthly":
        month_labels = list(available_months.keys())

        selected_months = st.multiselect(
            "📅 Select Reporting Months",
            options=month_labels,
            default=[month_labels[-1]] if month_labels else [],
            key="selected_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            {
                "label": month,
                "vrm": available_months[month]["vrm"],
                "drm": available_months[month]["drm"],
            }
            for month in selected_months
        ]

        period_label = ", ".join(selected_months)

    elif reporting_type == "Quarterly":
        available_quarters = discover_quarterly_files(DATA_DIR)

        if not available_quarters:
            st.error("⚠️ No quarterly data available.")
            return

        quarter_labels = list(available_quarters.keys())

        selected_quarter = st.selectbox(
            "📅 Reporting Quarter",
            options=quarter_labels,
            index=len(quarter_labels) - 1,
            key="selected_quarter",
        )

        quarter_entries = available_quarters[selected_quarter]

        month_options = [
            item["label"]
            for item in quarter_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_quarter_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in quarter_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_quarter} — "
            + ", ".join(selected_months)
        )

    else:
        available_half_years = discover_half_yearly_files(DATA_DIR)

        if not available_half_years:
            st.error("⚠️ No half-yearly data available.")
            return

        hy_labels = list(available_half_years.keys())

        selected_hy = st.selectbox(
            "📅 Reporting Half-Year",
            options=hy_labels,
            index=len(hy_labels) - 1,
            key="selected_half_year",
        )

        hy_entries = available_half_years[selected_hy]

        month_options = [
            item["label"]
            for item in hy_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_hy_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in hy_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_hy} — "
            + ", ".join(selected_months)
        )

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {period_label}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD VRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading VRM data for {period_label}…"):
        data = load_selected_vrm_data(selected_entries)

    if not data or data["active"].empty:
        st.error(f"⚠️ Could not load VRM data for {period_label}.")
        return

    df_raw  = data["active"]
    dropped = data["dropped"]
    new_reg = data["new_reg"]

    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD DRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading DRM data for {period_label}…"):
        drm_data_raw = load_selected_drm_data(selected_entries)

    # ── Build sidebar — content switches based on active_tab flag ─────────────
    with st.sidebar:
        st.markdown("---")

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

        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())
            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

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

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

                _dmask = pd.Series([True] * len(_sr), index=_sr.index)
                if sel_drm_donors:
                    _dmask &= _sr["Donor"].isin(sel_drm_donors)
                if sel_drm_states:
                    _dmask &= _sr["State"].isin(sel_states)
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

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
        new_vols            = new_reg["User ID"].nunique() if "User ID" in new_reg.columns else 0
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
        if "Date Joined" in new_reg.columns and "User ID" in new_reg.columns:
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
        if "Gender" in new_reg.columns and "User ID" in new_reg.columns:
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
            st.info("🔍 Gender breakdown (En Boys / En Girls) isn't available in this export.")

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
    # TAB 4 — DRM CLIENT REPORT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:
        if not drm_data_raw or "sess" not in drm_data_raw or drm_data_raw["sess"].empty:
            st.error(f"⚠️ No DRM file found or DRM data is empty for **{period_label}**.")
            st.stop()

        sess_raw = drm_data_raw["sess"]
        drm_raw  = drm_data_raw["drm"]
        od_raw   = drm_data_raw["od"]

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

        if sess.empty or drm.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.info(
                "🔍 **No data available for the selected filters.**\n\n"
                "The combination of Donor, State, and Subject you selected returned no sessions "
                "or no active centres. Try broadening your selection in the sidebar."
            )
            st.stop()

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
                            &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {period_label}
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

        total_sessions = len(sess)
        comp_sessions  = int((sess["Session_status"] == "Completed").sum())
        off_sessions   = int((sess["Session_status"] == "Offline").sum())
        canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
        total_clh_drm  = int(sess["Present/CLH"].sum())
        total_students = int(drm["Enrolled"].sum())
        avg_att_drm    = sess["Attendance%"].mean()
        total_centres  = drm["Center Name"].nunique()
        drm_total_vols = sess["Volunteer_id"].nunique() if "Volunteer_id" in sess.columns else 0
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

        for col_w, lbl, val, color, sub in [
            (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
            (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique active"),
            (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned in period"),
            (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
            (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

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
            (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{(off_sessions/total_sessions*100) if total_sessions else 0:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")

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

        st.markdown("#### ❌ Cancellations & CLH Split")

        col_ca1, col_ca2 = st.columns([1, 1.4])

        with col_ca1:
            st.markdown("##### Cancellation reasons")
            if "Cancel_reason" in sess.columns:
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
            else:
                st.info("Cancellation reasons are not available in this dataset.")

        with col_ca2:
            st.markdown("##### State-wise CLH split")
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

        st.markdown("#### 📚 Subject & Grade Analytics")

        col_sg1, col_sg2 = st.columns(2)

        with col_sg1:
            st.markdown("##### Sessions by subject & grade")
            if "Grade" in sess.columns and "Session_id" in sess.columns:
                subj_grade_sess = (
                    sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                    .count().reset_index()
                    .rename(columns={"Session_id": "Sessions"})
                )
                pivot_sgs = (
                    subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                    .fillna(0)
                )
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

        with col_sg2:
            st.markdown("##### Session completion % — subject & grade")
            if "Grade" in sess.columns and "Session_id" in sess.columns:
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
                if 'pivot_sgs' in locals():
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

        st.markdown("#### ⏱️ CLH by Subject & Grade")

        col_clh1, col_clh2 = st.columns([1.4, 1])

        with col_clh1:
            st.markdown("##### CLH heatmap — subject × grade")
            if "Grade" in sess.columns:
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

        st.markdown("#### 🙋 Volunteer Performance Scorecard")

        if {"Volunteer_id", "Teacher_name", "Session_id"}.issubset(sess.columns):
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

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                    padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                eVidyaloka Programme Impact Report &nbsp;·&nbsp; {period_label}
                &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                {drm_total_vols} volunteers &nbsp;·&nbsp;
                {total_students:,} students &nbsp;·&nbsp;
                {total_clh_drm:,} child learning hours
            </div>""",
            unsafe_allow_html=True,
        )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — TOPIC & SUBTOPIC ANALYTICS (INTEGRATED TOPIC DASHBOARD)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("### 📚 Topic & SubTopic Analytics Dashboard")

        uploaded_topic_file = st.file_uploader(
            "📁 Upload Topic/SubTopic Excel File",
            type=["xlsx", "xls"],
            key="topic_file_uploader",
            help="Upload the Excel file containing 'Topic-SubTopic', 'Cancelled Sessions', and 'Offline Sessions' sheets."
        )

        if uploaded_topic_file is None:
            st.info("👆 Please upload a Topic/SubTopic Excel file above to activate this dashboard.")
        else:
            @st.cache_data
            def load_topic_uploaded_data(file):
                try:
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names

                    def _get_sheet_name(target):
                        norm_target = " ".join(str(target).strip().lower().split())
                        for sheet in sheets:
                            if " ".join(str(sheet).strip().lower().split()) == norm_target:
                                return sheet
                        return target

                    t_sheet = _get_sheet_name("Topic-SubTopic")
                    c_sheet = _get_sheet_name("Cancelled Sessions")
                    o_sheet = _get_sheet_name("Offline Sessions")

                    t_df = pd.read_excel(file, sheet_name=t_sheet) if t_sheet in sheets else pd.DataFrame()
                    c_df = pd.read_excel(file, sheet_name=c_sheet) if c_sheet in sheets else pd.DataFrame()
                    o_df = pd.read_excel(file, sheet_name=o_sheet) if o_sheet in sheets else pd.DataFrame()

                    return t_df, c_df, o_df
                except Exception as e:
                    st.error(f"Error reading uploaded file: {e}")
                    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            topic_df, cancelled_df, offline_df = load_topic_uploaded_data(uploaded_topic_file)

            if topic_df.empty:
                st.error("The uploaded file does not contain a valid 'Topic-SubTopic' sheet or has no data.")
            else:
                # ── IN-TAB FILTERS ──────────────────────────────────────────
                with st.expander("📊 Topic Filters", expanded=True):
                    fc1, fc2, fc3 = st.columns(3)

                    state_col   = "state" if "state" in topic_df.columns else "State"
                    subject_col = "subject" if "subject" in topic_df.columns else "Subject"
                    grade_col   = "grade" if "grade" in topic_df.columns else "Grade"

                    with fc1:
                        st_opts = sorted(topic_df[state_col].dropna().unique()) if state_col in topic_df.columns else []
                        t_state_filter = st.multiselect("State", st_opts, key="tab5_state")

                    state_df = topic_df[topic_df[state_col].isin(t_state_filter)] if t_state_filter and state_col in topic_df.columns else topic_df.copy()

                    with fc2:
                        sub_opts = sorted(state_df[subject_col].dropna().unique()) if subject_col in state_df.columns else []
                        t_subject_filter = st.multiselect("Subject", sub_opts, key="tab5_subject")

                    subject_df = state_df[state_df[subject_col].isin(t_subject_filter)] if t_subject_filter and subject_col in state_df.columns else state_df.copy()

                    with fc3:
                        grd_opts = sorted(subject_df[grade_col].dropna().unique()) if grade_col in subject_df.columns else []
                        t_grade_filter = st.multiselect("Grade", grd_opts, key="tab5_grade")

                    filtered_topic_df = subject_df[subject_df[grade_col].isin(t_grade_filter)] if t_grade_filter and grade_col in subject_df.columns else subject_df.copy()

                # ── KPI METRICS ──────────────────────────────────────────────
                total_sessions   = len(filtered_topic_df)
                unique_topics    = filtered_topic_df["topic_name"].nunique() if "topic_name" in filtered_topic_df.columns else 0
                unique_subtopics = filtered_topic_df["sub_topic_name"].nunique() if "sub_topic_name" in filtered_topic_df.columns else 0
                unique_centers   = filtered_topic_df["center_id"].nunique() if "center_id" in filtered_topic_df.columns else (
                    filtered_topic_df["center_name"].nunique() if "center_name" in filtered_topic_df.columns else 0
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Sessions", f"{total_sessions:,}")
                with col2:
                    st.metric("Unique Topics", f"{unique_topics:,}")
                with col3:
                    st.metric("Unique Subtopics", f"{unique_subtopics:,}")
                with col4:
                    st.metric("Active Centers", f"{unique_centers:,}")

                st.markdown("---")

                # ── DASHBOARD SUB-TABS ────────────────────────────────────────
                top_tab1, top_tab2, top_tab3 = st.tabs([
                    "📋 Topic Analysis",
                    "❌ Cancelled Sessions",
                    "⏸️ Offline Sessions"
                ])

                # 1. TOPIC ANALYSIS
                with top_tab1:
                    if "topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Topics")
                        topic_summary = (
                            filtered_topic_df
                            .groupby("topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig = px.bar(
                            topic_summary, x="Sessions", y="topic_name",
                            orientation="h", title="Top 20 Topics"
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    if "sub_topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Sub Topics")
                        subtopic_summary = (
                            filtered_topic_df
                            .groupby("sub_topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig2 = px.bar(
                            subtopic_summary, x="Sessions", y="sub_topic_name",
                            orientation="h", title="Top 20 Sub Topics"
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    if subject_col in filtered_topic_df.columns:
                        st.subheader("Subject Distribution")
                        subject_chart = (
                            filtered_topic_df
                            .groupby(subject_col)
                            .size()
                            .reset_index(name="Count")
                        )
                        fig3 = px.pie(
                            subject_chart, names=subject_col, values="Count"
                        )
                        st.plotly_chart(fig3, use_container_width=True)

                # 2. CANCELLED SESSIONS
                with top_tab2:
                    st.subheader("Cancelled Session Analysis")
                    if not cancelled_df.empty:
                        st.metric("Total Cancelled Sessions", f"{len(cancelled_df):,}")

                        cancel_reason_col = (
                            "cancel_reason" if "cancel_reason" in cancelled_df.columns else (
                                "Cancel_reason" if "Cancel_reason" in cancelled_df.columns else None
                            )
                        )
                        if cancel_reason_col:
                            cancel_chart = (
                                cancelled_df
                                .groupby(cancel_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                                .head(15)
                            )
                            fig4 = px.bar(
                                cancel_chart, x="Count", y=cancel_reason_col,
                                orientation="h", title="Cancellation Reasons"
                            )
                            st.plotly_chart(fig4, use_container_width=True)

                        canc_state_col = "State" if "State" in cancelled_df.columns else (
                            "state" if "state" in cancelled_df.columns else None
                        )
                        if canc_state_col:
                            state_cancel = (
                                cancelled_df
                                .groupby(canc_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig5 = px.bar(
                                state_cancel, x=canc_state_col, y="Count",
                                title="Cancelled Sessions by State"
                            )
                            st.plotly_chart(fig5, use_container_width=True)
                    else:
                        st.info("No cancelled sessions data found in the uploaded file.")

                # 3. OFFLINE SESSIONS
                with top_tab3:
                    st.subheader("Offline Session Analysis")
                    if not offline_df.empty:
                        st.metric("Total Offline Sessions", f"{len(offline_df):,}")

                        offline_reason_col = (
                            "Offline reason" if "Offline reason" in offline_df.columns else (
                                "offline_reason" if "offline_reason" in offline_df.columns else None
                            )
                        )
                        if offline_reason_col:
                            offline_reason = (
                                offline_df
                                .groupby(offline_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig6 = px.bar(
                                offline_reason, x="Count", y=offline_reason_col,
                                orientation="h", title="Offline Reasons"
                            )
                            st.plotly_chart(fig6, use_container_width=True)

                        off_state_col = "State" if "State" in offline_df.columns else (
                            "state" if "state" in offline_df.columns else None
                        )
                        if off_state_col:
                            state_offline = (
                                offline_df
                                .groupby(off_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig7 = px.bar(
                                state_offline, x=off_state_col, y="Count",
                                title="Offline Sessions by State"
                            )
                            st.plotly_chart(fig7, use_container_width=True)
                    else:
                        st.info("No offline sessions data found in the uploaded file.")

                # ── EXPORT DATA ──────────────────────────────────────────────
                st.markdown("---")
                st.download_button(
                    "📥 Download Filtered Topic Data",
                    filtered_topic_df.to_csv(index=False),
                    file_name=f"filtered_{uploaded_topic_file.name.replace('.xlsx', '.csv')}",
                    mime="text/csv",
                    key="download_topic_csv"
                )
in above code can you tell me how the avg attendance % is being calculated like from which sheet like VRM OR DRM attendance is being considered

In the code, the calculation of Average Attendance % depends on which tab or section of the dashboard you are looking at:

1. Volunteer Overview KPIs (Top Metric Cards)
In the main overview KPI section at the top, Avg Attendance is calculated directly from the VRM file:

Source Sheet: Active VT sheet from the VRM workbook.

Calculation: It computes the simple average across all rows in the active volunteer offerings matching your sidebar filters:

Python
avg_att = df["Attendance%"].mean()
2. Academic Health Tab (Tab 3)
In Tab 3 ("📚 Academic Health"), the attendance figures and heatmaps are also derived from the VRM file:

Source Sheet: Active VT sheet from the VRM workbook.

State × Subject Heatmap:

Python
heat_df = df.groupby(["State", "Subject_clean"])["Attendance%"].mean().reset_index()
Attendance Histogram: Plotly histogram mapping df["Attendance%"] from the VRM data.

3. DRM Client Report Tab (Tab 4)
In Tab 4 ("📊 DRM Client Report"), all average attendance calculations switch completely to the DRM file:

Source Sheet: SESSION DUMP sheet from the DRM workbook.

Programme Snapshot KPI:

Python
avg_att_drm = sess["Attendance%"].mean()
State-wise Avg Attendance Chart:

Python
state_sess = sess.groupby("State").agg(
    Attendance=("Attendance%", "mean")
)
Volunteer Performance Scorecard:

Python
vol_perf = sess.groupby(["Volunteer_id", "Teacher_name"]).agg(
    Avg_Att=("Attendance%", "mean")
)
Summary Table
Dashboard Area / Tab	Source Workbook	Source Sheet	Column Name
Top KPI Cards	VRM	Active VT	Attendance%
Tab 3: Academic Health	VRM	Active VT	Attendance%
Tab 4: DRM Client Report	DRM	SESSION DUMP	Attendance%
Tab 4: Full Centre Scorecard	DRM	DRM	Attendance %
UnboundLocalError: local variable 'sel_states' referenced before assignment

Traceback:

File "/var/streamlitdashboard/jstreamlit/Impact-Analytics-Dashboard/app.py", line 289, in <module>

render_ops_dashboard()

File "/var/streamlitdashboard/jstreamlit/Impact-Analytics-Dashboard/ops_dashboard.py", line 923, in render_ops_dashboard

_dmask &= _sr["State"].isin(sel_states) .when clicked on DRM filters shows above error

The error occurs because sel_states (the state selection variable for VRM filters) hasn't been defined yet when you switch to the DRM Filters tab.

Line 923 contains a small typo where sel_states was used instead of _drm_linked_states:

Python
# ❌ Problematic line in the DRM Filters block:
_dmask &= _sr["State"].isin(sel_states)
Fix
In ops_dashboard.py, locate line 923 inside the DRM sidebar filter block:

Python
# Change this:
_dmask &= _sr["State"].isin(sel_states)

# To this:
_dmask &= _sr["State"].isin(_drm_linked_states)
Corrected Sidebar Section Code
Here is the updated with st.sidebar: code block with the fix applied:

Python
        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

                _dmask = pd.Series([True] * len(_sr), index=_sr.index)
                if sel_drm_donors:
                    _dmask &= _sr["Donor"].isin(sel_drm_donors)
                if sel_drm_states:
                    # FIX APPLIED HERE: replaced sel_states with sel_drm_states (or _drm_linked_states)
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())
i updated the code with this # To this:

_dmask &= _sr["State"].isin(_drm_linked_states) and what to do

You should use sel_drm_states instead of _drm_linked_states.

If you use _drm_linked_states, filtering by individual states in the DRM filter dropdown won't actually filter the subjects because it reads from all available linked states rather than what the user selected in the dropdown.

Here is the exact fix:

1. Update line 923 to:
Python
_dmask &= _sr["State"].isin(sel_drm_states)
2. Complete Fixed DRM Sidebar Block
For reference, here is how the entire else: block inside with st.sidebar: should look:

Python
        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())
Once you save ops_dashboard.py with sel_drm_states, the UnboundLocalError will be resolved and state-filtering in the DRM tab will work.

where should i paste this block which lineComplete Fixed DRM Sidebar Block

In ops_dashboard.py, locate the else: block inside with st.sidebar: (around line 900) and replace that entire else: block with this complete code:

Python
        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())
Context for where it sits in the file:
It goes directly inside the with st.sidebar: section, right after the if st.session_state.active_tab == "vrm": block finishes:

Python
        if st.session_state.active_tab == "vrm":
            # ... VRM Filters code ...
            
        else:
            # 👈 PASTE THE ENTIRE BLOCK HERE
can you share the complete updated ops.dashbord code file

Here is the complete, fully updated ops_dashboard.py file with the Half-Yearly radio button option, Tab 5 ("Topic & SubTopic Analytics"), and the sidebar filter bug fix (sel_drm_states) all fully integrated.

Python
"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,
Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().

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
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

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
    them by reporting month.
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


# ─────────────────────────────────────────────────────────────────────────────
# QUARTERLY & HALF-YEARLY REPORTING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_quarter_label(month_num: int, year: int) -> str:
    """Financial year quarters:
       Q1 = Apr-Jun
       Q2 = Jul-Sep
       Q3 = Oct-Dec
       Q4 = Jan-Mar
    """
    if month_num in (4, 5, 6):
        return f"Q1 {year} (Apr–Jun)"
    elif month_num in (7, 8, 9):
        return f"Q2 {year} (Jul–Sep)"
    elif month_num in (10, 11, 12):
        return f"Q3 {year} (Oct–Dec)"
    else:
        return f"Q4 {year - 1}-{year} (Jan–Mar)"


def discover_quarterly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial quarters."""
    monthly = discover_monthly_files(data_dir)
    quarters = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        quarter = get_quarter_label(month_num, year)
        if quarter not in quarters:
            quarters[quarter] = []

        quarters[quarter].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for quarter in quarters:
        quarters[quarter] = sorted(
            quarters[quarter],
            key=lambda x: (x["year"], x["month"])
        )

    return quarters


def get_half_yearly_label(month_num: int, year: int) -> str:
    """Financial year half-years:
       H1 = Apr-Sep
       H2 = Oct-Mar
    """
    if month_num in (4, 5, 6, 7, 8, 9):
        return f"H1 {year} (Apr–Sep)"
    else:
        fy_year = year if month_num in (10, 11, 12) else year - 1
        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"


def discover_half_yearly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial half-years."""
    monthly = discover_monthly_files(data_dir)
    half_years = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        hy_label = get_half_yearly_label(month_num, year)
        if hy_label not in half_years:
            half_years[hy_label] = []

        half_years[hy_label].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for hy in half_years:
        half_years[hy] = sorted(
            half_years[hy],
            key=lambda x: (x["year"], x["month"])
        )

    return half_years


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-MONTH VRM AND DRM COMBINE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_selected_vrm_data(selected_entries: list) -> dict:
    """
    Loads VRM data for multiple selected month entries and concatenates 
    them together across 'active', 'dropped', and 'new_reg' sheets.
    """
    active_dfs, dropped_dfs, new_reg_dfs = [], [], []

    for entry in selected_entries:
        vrm_path = entry.get("vrm")
        if vrm_path and os.path.exists(vrm_path):
            month_data = load_data(vrm_path)
            if month_data:
                if "active" in month_data and not month_data["active"].empty:
                    df_act = month_data["active"].copy()
                    df_act["Reporting Period"] = entry["label"]
                    active_dfs.append(df_act)
                
                if "dropped" in month_data and not month_data["dropped"].empty:
                    df_drp = month_data["dropped"].copy()
                    df_drp["Reporting Period"] = entry["label"]
                    dropped_dfs.append(df_drp)
                
                if "new_reg" in month_data and not month_data["new_reg"].empty:
                    df_nr = month_data["new_reg"].copy()
                    df_nr["Reporting Period"] = entry["label"]
                    new_reg_dfs.append(df_nr)

    return {
        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),
        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),
        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),
    }


def load_selected_drm_data(selected_entries: list) -> dict:
    """
    Loads DRM data for multiple selected month entries and concatenates 
    them across all individual DRM sheets.
    """
    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []

    for entry in selected_entries:
        drm_path = entry.get("drm")
        if drm_path and os.path.exists(drm_path):
            m_drm = load_drm_data(drm_path)
            if m_drm:
                for key, container in [
                    ("sess", sess_dfs),
                    ("drm", drm_dfs),
                    ("ac", ac_dfs),
                    ("od", od_dfs),
                    ("ne", ne_dfs),
                ]:
                    if key in m_drm and not m_drm[key].empty:
                        df_sheet = m_drm[key].copy()
                        df_sheet["Reporting Period"] = entry["label"]
                        container.append(df_sheet)

    return {
        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),
        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),
        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),
        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),
        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),
    }


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

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

    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))
    dropped.columns = dropped.columns.str.strip()
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

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

DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

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

    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))
    ac.columns = ac.columns.str.strip()

    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

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

    # ─────────────────────────────────────────────────────────────────────────
    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)
    # ─────────────────────────────────────────────────────────────────────────

    reporting_type = st.radio(
        "📊 Reporting Type",
        ["Monthly", "Quarterly", "Half-Yearly"],
        horizontal=True,
        key="reporting_type",
    )

    selected_entries = []
    period_label = ""

    if reporting_type == "Monthly":
        month_labels = list(available_months.keys())

        selected_months = st.multiselect(
            "📅 Select Reporting Months",
            options=month_labels,
            default=[month_labels[-1]] if month_labels else [],
            key="selected_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            {
                "label": month,
                "vrm": available_months[month]["vrm"],
                "drm": available_months[month]["drm"],
            }
            for month in selected_months
        ]

        period_label = ", ".join(selected_months)

    elif reporting_type == "Quarterly":
        available_quarters = discover_quarterly_files(DATA_DIR)

        if not available_quarters:
            st.error("⚠️ No quarterly data available.")
            return

        quarter_labels = list(available_quarters.keys())

        selected_quarter = st.selectbox(
            "📅 Reporting Quarter",
            options=quarter_labels,
            index=len(quarter_labels) - 1,
            key="selected_quarter",
        )

        quarter_entries = available_quarters[selected_quarter]

        month_options = [
            item["label"]
            for item in quarter_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_quarter_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in quarter_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_quarter} — "
            + ", ".join(selected_months)
        )

    else:
        available_half_years = discover_half_yearly_files(DATA_DIR)

        if not available_half_years:
            st.error("⚠️ No half-yearly data available.")
            return

        hy_labels = list(available_half_years.keys())

        selected_hy = st.selectbox(
            "📅 Reporting Half-Year",
            options=hy_labels,
            index=len(hy_labels) - 1,
            key="selected_half_year",
        )

        hy_entries = available_half_years[selected_hy]

        month_options = [
            item["label"]
            for item in hy_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_hy_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in hy_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_hy} — "
            + ", ".join(selected_months)
        )

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {period_label}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD VRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading VRM data for {period_label}…"):
        data = load_selected_vrm_data(selected_entries)

    if not data or data["active"].empty:
        st.error(f"⚠️ Could not load VRM data for {period_label}.")
        return

    df_raw  = data["active"]
    dropped = data["dropped"]
    new_reg = data["new_reg"]

    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD DRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading DRM data for {period_label}…"):
        drm_data_raw = load_selected_drm_data(selected_entries)

    # ── Build sidebar — content switches based on active_tab flag ─────────────
    with st.sidebar:
        st.markdown("---")

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

        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())
            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

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

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

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
        new_vols            = new_reg["User ID"].nunique() if "User ID" in new_reg.columns else 0
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
        if "Date Joined" in new_reg.columns and "User ID" in new_reg.columns:
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
        if "Gender" in new_reg.columns and "User ID" in new_reg.columns:
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
            st.info("🔍 Gender breakdown (En Boys / En Girls) isn't available in this export.")

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
    # TAB 4 — DRM CLIENT REPORT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:
        if not drm_data_raw or "sess" not in drm_data_raw or drm_data_raw["sess"].empty:
            st.error(f"⚠️ No DRM file found or DRM data is empty for **{period_label}**.")
            st.stop()

        sess_raw = drm_data_raw["sess"]
        drm_raw  = drm_data_raw["drm"]
        od_raw   = drm_data_raw["od"]

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

        if sess.empty or drm.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.info(
                "🔍 **No data available for the selected filters.**\n\n"
                "The combination of Donor, State, and Subject you selected returned no sessions "
                "or no active centres. Try broadening your selection in the sidebar."
            )
            st.stop()

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
                            &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {period_label}
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

        total_sessions = len(sess)
        comp_sessions  = int((sess["Session_status"] == "Completed").sum())
        off_sessions   = int((sess["Session_status"] == "Offline").sum())
        canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
        total_clh_drm  = int(sess["Present/CLH"].sum())
        total_students = int(drm["Enrolled"].sum())
        avg_att_drm    = sess["Attendance%"].mean()
        total_centres  = drm["Center Name"].nunique()
        drm_total_vols = sess["Volunteer_id"].nunique() if "Volunteer_id" in sess.columns else 0
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

        for col_w, lbl, val, color, sub in [
            (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
            (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique active"),
            (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned in period"),
            (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
            (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

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
            (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{(off_sessions/total_sessions*100) if total_sessions else 0:.1f}% of planned"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")

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

        st.markdown("#### ❌ Cancellations & CLH Split")

        col_ca1, col_ca2 = st.columns([1, 1.4])

        with col_ca1:
            st.markdown("##### Cancellation reasons")
            if "Cancel_reason" in sess.columns:
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
            else:
                st.info("Cancellation reasons are not available in this dataset.")

        with col_ca2:
            st.markdown("##### State-wise CLH split")
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

        st.markdown("#### 📚 Subject & Grade Analytics")

        col_sg1, col_sg2 = st.columns(2)

        with col_sg1:
            st.markdown("##### Sessions by subject & grade")
            if "Grade" in sess.columns and "Session_id" in sess.columns:
                subj_grade_sess = (
                    sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                    .count().reset_index()
                    .rename(columns={"Session_id": "Sessions"})
                )
                pivot_sgs = (
                    subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                    .fillna(0)
                )
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

        with col_sg2:
            st.markdown("##### Session completion % — subject & grade")
            if "Grade" in sess.columns and "Session_id" in sess.columns:
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
                if 'pivot_sgs' in locals():
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

        st.markdown("#### ⏱️ CLH by Subject & Grade")

        col_clh1, col_clh2 = st.columns([1.4, 1])

        with col_clh1:
            st.markdown("##### CLH heatmap — subject × grade")
            if "Grade" in sess.columns:
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

        st.markdown("#### 🙋 Volunteer Performance Scorecard")

        if {"Volunteer_id", "Teacher_name", "Session_id"}.issubset(sess.columns):
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

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                    padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                eVidyaloka Programme Impact Report &nbsp;·&nbsp; {period_label}
                &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                {drm_total_vols} volunteers &nbsp;·&nbsp;
                {total_students:,} students &nbsp;·&nbsp;
                {total_clh_drm:,} child learning hours
            </div>""",
            unsafe_allow_html=True,
        )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — TOPIC & SUBTOPIC ANALYTICS (INTEGRATED TOPIC DASHBOARD)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("### 📚 Topic & SubTopic Analytics Dashboard")

        uploaded_topic_file = st.file_uploader(
            "📁 Upload Topic/SubTopic Excel File",
            type=["xlsx", "xls"],
            key="topic_file_uploader",
            help="Upload the Excel file containing 'Topic-SubTopic', 'Cancelled Sessions', and 'Offline Sessions' sheets."
        )

        if uploaded_topic_file is None:
            st.info("👆 Please upload a Topic/SubTopic Excel file above to activate this dashboard.")
        else:
            @st.cache_data
            def load_topic_uploaded_data(file):
                try:
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names

                    def _get_sheet_name(target):
                        norm_target = " ".join(str(target).strip().lower().split())
                        for sheet in sheets:
                            if " ".join(str(sheet).strip().lower().split()) == norm_target:
                                return sheet
                        return target

                    t_sheet = _get_sheet_name("Topic-SubTopic")
                    c_sheet = _get_sheet_name("Cancelled Sessions")
                    o_sheet = _get_sheet_name("Offline Sessions")

                    t_df = pd.read_excel(file, sheet_name=t_sheet) if t_sheet in sheets else pd.DataFrame()
                    c_df = pd.read_excel(file, sheet_name=c_sheet) if c_sheet in sheets else pd.DataFrame()
                    o_df = pd.read_excel(file, sheet_name=o_sheet) if o_sheet in sheets else pd.DataFrame()

                    return t_df, c_df, o_df
                except Exception as e:
                    st.error(f"Error reading uploaded file: {e}")
                    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            topic_df, cancelled_df, offline_df = load_topic_uploaded_data(uploaded_topic_file)

            if topic_df.empty:
                st.error("The uploaded file does not contain a valid 'Topic-SubTopic' sheet or has no data.")
            else:
                # ── IN-TAB FILTERS ──────────────────────────────────────────
                with st.expander("📊 Topic Filters", expanded=True):
                    fc1, fc2, fc3 = st.columns(3)

                    state_col   = "state" if "state" in topic_df.columns else "State"
                    subject_col = "subject" if "subject" in topic_df.columns else "Subject"
                    grade_col   = "grade" if "grade" in topic_df.columns else "Grade"

                    with fc1:
                        st_opts = sorted(topic_df[state_col].dropna().unique()) if state_col in topic_df.columns else []
                        t_state_filter = st.multiselect("State", st_opts, key="tab5_state")

                    state_df = topic_df[topic_df[state_col].isin(t_state_filter)] if t_state_filter and state_col in topic_df.columns else topic_df.copy()

                    with fc2:
                        sub_opts = sorted(state_df[subject_col].dropna().unique()) if subject_col in state_df.columns else []
                        t_subject_filter = st.multiselect("Subject", sub_opts, key="tab5_subject")

                    subject_df = state_df[state_df[subject_col].isin(t_subject_filter)] if t_subject_filter and subject_col in state_df.columns else state_df.copy()

                    with fc3:
                        grd_opts = sorted(subject_df[grade_col].dropna().unique()) if grade_col in subject_df.columns else []
                        t_grade_filter = st.multiselect("Grade", grd_opts, key="tab5_grade")

                    filtered_topic_df = subject_df[subject_df[grade_col].isin(t_grade_filter)] if t_grade_filter and grade_col in subject_df.columns else subject_df.copy()

                # ── KPI METRICS ──────────────────────────────────────────────
                total_sessions   = len(filtered_topic_df)
                unique_topics    = filtered_topic_df["topic_name"].nunique() if "topic_name" in filtered_topic_df.columns else 0
                unique_subtopics = filtered_topic_df["sub_topic_name"].nunique() if "sub_topic_name" in filtered_topic_df.columns else 0
                unique_centers   = filtered_topic_df["center_id"].nunique() if "center_id" in filtered_topic_df.columns else (
                    filtered_topic_df["center_name"].nunique() if "center_name" in filtered_topic_df.columns else 0
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Sessions", f"{total_sessions:,}")
                with col2:
                    st.metric("Unique Topics", f"{unique_topics:,}")
                with col3:
                    st.metric("Unique Subtopics", f"{unique_subtopics:,}")
                with col4:
                    st.metric("Active Centers", f"{unique_centers:,}")

                st.markdown("---")

                # ── DASHBOARD SUB-TABS ────────────────────────────────────────
                top_tab1, top_tab2, top_tab3 = st.tabs([
                    "📋 Topic Analysis",
                    "❌ Cancelled Sessions",
                    "⏸️ Offline Sessions"
                ])

                # 1. TOPIC ANALYSIS
                with top_tab1:
                    if "topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Topics")
                        topic_summary = (
                            filtered_topic_df
                            .groupby("topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig = px.bar(
                            topic_summary, x="Sessions", y="topic_name",
                            orientation="h", title="Top 20 Topics"
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    if "sub_topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Sub Topics")
                        subtopic_summary = (
                            filtered_topic_df
                            .groupby("sub_topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig2 = px.bar(
                            subtopic_summary, x="Sessions", y="sub_topic_name",
                            orientation="h", title="Top 20 Sub Topics"
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    if subject_col in filtered_topic_df.columns:
                        st.subheader("Subject Distribution")
                        subject_chart = (
                            filtered_topic_df
                            .groupby(subject_col)
                            .size()
                            .reset_index(name="Count")
                        )
                        fig3 = px.pie(
                            subject_chart, names=subject_col, values="Count"
                        )
                        st.plotly_chart(fig3, use_container_width=True)

                # 2. CANCELLED SESSIONS
                with top_tab2:
                    st.subheader("Cancelled Session Analysis")
                    if not cancelled_df.empty:
                        st.metric("Total Cancelled Sessions", f"{len(cancelled_df):,}")

                        cancel_reason_col = (
                            "cancel_reason" if "cancel_reason" in cancelled_df.columns else (
                                "Cancel_reason" if "Cancel_reason" in cancelled_df.columns else None
                            )
                        )
                        if cancel_reason_col:
                            cancel_chart = (
                                cancelled_df
                                .groupby(cancel_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                                .head(15)
                            )
                            fig4 = px.bar(
                                cancel_chart, x="Count", y=cancel_reason_col,
                                orientation="h", title="Cancellation Reasons"
                            )
                            st.plotly_chart(fig4, use_container_width=True)

                        canc_state_col = "State" if "State" in cancelled_df.columns else (
                            "state" if "state" in cancelled_df.columns else None
                        )
                        if canc_state_col:
                            state_cancel = (
                                cancelled_df
                                .groupby(canc_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig5 = px.bar(
                                state_cancel, x=canc_state_col, y="Count",
                                title="Cancelled Sessions by State"
                            )
                            st.plotly_chart(fig5, use_container_width=True)
                    else:
                        st.info("No cancelled sessions data found in the uploaded file.")

                # 3. OFFLINE SESSIONS
                with top_tab3:
                    st.subheader("Offline Session Analysis")
                    if not offline_df.empty:
                        st.metric("Total Offline Sessions", f"{len(offline_df):,}")

                        offline_reason_col = (
                            "Offline reason" if "Offline reason" in offline_df.columns else (
                                "offline_reason" if "offline_reason" in offline_df.columns else None
                            )
                        )
                        if offline_reason_col:
                            offline_reason = (
                                offline_df
                                .groupby(offline_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig6 = px.bar(
                                offline_reason, x="Count", y=offline_reason_col,
                                orientation="h", title="Offline Reasons"
                            )
                            st.plotly_chart(fig6, use_container_width=True)

                        off_state_col = "State" if "State" in offline_df.columns else (
                            "state" if "state" in offline_df.columns else None
                        )
                        if off_state_col:
                            state_offline = (
                                offline_df
                                .groupby(off_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig7 = px.bar(
                                state_offline, x=off_state_col, y="Count",
                                title="Offline Sessions by State"
                            )
                            st.plotly_chart(fig7, use_container_width=True)
                    else:
                        st.info("No offline sessions data found in the uploaded file.")

                # ── EXPORT DATA ──────────────────────────────────────────────
                st.markdown("---")
                st.download_button(
                    "📥 Download Filtered Topic Data",
                    filtered_topic_df.to_csv(index=False),
                    file_name=f"filtered_{uploaded_topic_file.name.replace('.xlsx', '.csv')}",
                    mime="text/csv",
                    key="download_topic_csv"
                )
File "/mount/src/impact_dashboard/app.py", line 224, in <module>

render_ops_dashboard()

~~~~~~~~~~~~~~~~~~~~^^

File "/mount/src/impact_dashboard/ops_dashboard.py", line 786, in render_ops_dashboard

data = load_selected_vrm_data(selected_entries)

File "/mount/src/impact_dashboard/ops_dashboard.py", line 202, in load_selected_vrm_data

month_data = load_data(vrm_path)

File "/mount/src/impact_dashboard/ops_dashboard.py", line 460, in load_data

dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))

File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_base.py", line 526, in read_excel

data = io._reader.parse(

sheet_name=sheet_name,

...<20 lines>...

dtype_backend=dtype_backend,

)

File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_base.py", line 787, in parse

sheet = self.get_sheet_by_name(asheetname)

File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_openpyxl.py", line 706, in get_sheet_by_name

self.raise_if_bad_sheet_by_name(name)

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~^^^^^^

File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_base.py", line 637, in raise_if_bad_sheet_by_name

raise ValueError(f"Worksheet named '{name}' not found")

The error occurs because one of the uploaded/scanned VRM Excel files is missing the "Dropped VT" sheet (or the sheet name is spelled differently, e.g., missing spaces or having different casing).

Currently, pd.read_excel() is called directly without checking if the target sheet actually exists in xl.sheet_names.

Fix
In ops_dashboard.py, update load_data() and load_drm_data() to check whether a sheet exists in xl.sheet_names before attempting to read it, returning an empty DataFrame if it is missing.

Here are the updated data loading functions:

Python
# ─────────────────────────────────────────────────────────────────────────────
# VRM DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)
    all_sheets = xl.sheet_names

    # Load Active VT
    active_sheet_name = _find_sheet(xl, SHEET_ACTIVE)
    if active_sheet_name in all_sheets:
        active = pd.read_excel(xl, sheet_name=active_sheet_name)
    else:
        active = pd.DataFrame()

    if not active.empty:
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

    # Load Dropped VT safely
    dropped_sheet_name = _find_sheet(xl, SHEET_DROPPED)
    if dropped_sheet_name in all_sheets:
        dropped = pd.read_excel(xl, sheet_name=dropped_sheet_name)
        dropped.columns = dropped.columns.str.strip()
        if "State" in dropped.columns:
            dropped["State"] = (
                dropped["State"].astype(str)
                .str.replace("\u00ac\u00a0", " ", regex=False)
                .str.replace("\u00a0", " ", regex=False)
                .str.strip()
            )
    else:
        dropped = pd.DataFrame()

    # Load Newly Registered VT safely
    new_reg_sheet_name = _find_sheet(xl, SHEET_NEW_REG)
    if new_reg_sheet_name in all_sheets:
        new_reg = pd.read_excel(xl, sheet_name=new_reg_sheet_name)
        new_reg.columns = new_reg.columns.str.strip()
        if "Date Joined" in new_reg.columns:
            new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
        if "Reference" in new_reg.columns:
            new_reg["Ref_group"] = new_reg["Reference"].apply(_group_ref)
        if "Gender" in new_reg.columns:
            new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()
    else:
        new_reg = pd.DataFrame()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}
Improvements
Prevents Crashes: If any file lacks the "Dropped VT" or "Newly Registered VT " sheet, the code defaults to an empty DataFrame instead of raising ValueError.

Maintains Backwards Compatibility: If the sheet is present, it will load normally.

does above code has topic subtopic tab

Yes, the complete code shared earlier includes Tab 5: 📚 Topic & SubTopic Analytics.

Here is where it is defined and how it functions in that code:

1. Tab Declaration
In the tab setup section of render_ops_dashboard(), 5 tabs are created:

Python
tab_vol, tab_ctr, tab_aca, tab_drm, tab_top = st.tabs([
    "🙋 Volunteers",
    "🏫 Centres",
    "📚 Academic Health",
    "📊 DRM Client Report",
    "📚 Topic & SubTopic Analytics",
])
2. Tab 5 Implementation Highlights (with tab_top:)
Interactive File Uploader: Allows uploading a Topic/SubTopic Excel workbook directly within the dashboard.

Cascading In-Tab Filters: Interactive multiselect filters for State, Subject, and Grade.

Key Metric Summary: Displays total sessions, unique topics, unique subtopics, and active centers.

Sub-Tabs & Visualizations:

Topic Analysis: Bar charts for Top 20 Topics, Top 20 SubTopics, and a Subject Distribution pie chart.

Cancelled Sessions: Breakdown of cancellation reasons and state-wise totals.

Offline Sessions: Visual breakdown of offline session reasons and state-wise distributions.

Data Export: Provides a button to download the filtered topic dataset as a CSV.

but when we uploaded the code its not showing tab 5

The reason Tab 5 ("📚 Topic & SubTopic Analytics") is not displaying is due to how Streamlit evaluates the tab code block in ops_dashboard.py.

In your script, the data loading check for the VRM file (if not data or data["active"].empty:) happens before the st.tabs() declaration. If the VRM file fails to load or has an issue with a sheet, Streamlit triggers st.error() and stops execution before reaching the st.tabs() line, preventing Tab 5 (and all other tabs) from rendering.

The Fix
To ensure Tab 5 always renders reliably—even if VRM or DRM data has issues—you need to apply two key fixes:

Safe Excel Loading: Prevent missing/renamed sheets (Active VT, Dropped VT, Newly Registered VT ) from throwing errors.

Tab Layout Scaffolding: Move the st.tabs() declaration above data failure blocks or allow empty data frames to pass through safely.

Corrected Code (ops_dashboard.py)
Here is the complete code with safe sheet parsing and guaranteed rendering for all 5 tabs:

Python
"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,
Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().

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
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

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
    them by reporting month.
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


# ─────────────────────────────────────────────────────────────────────────────
# QUARTERLY & HALF-YEARLY REPORTING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_quarter_label(month_num: int, year: int) -> str:
    if month_num in (4, 5, 6):
        return f"Q1 {year} (Apr–Jun)"
    elif month_num in (7, 8, 9):
        return f"Q2 {year} (Jul–Sep)"
    elif month_num in (10, 11, 12):
        return f"Q3 {year} (Oct–Dec)"
    else:
        return f"Q4 {year - 1}-{year} (Jan–Mar)"


def discover_quarterly_files(data_dir: str) -> dict:
    monthly = discover_monthly_files(data_dir)
    quarters = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        quarter = get_quarter_label(month_num, year)
        if quarter not in quarters:
            quarters[quarter] = []

        quarters[quarter].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for quarter in quarters:
        quarters[quarter] = sorted(
            quarters[quarter],
            key=lambda x: (x["year"], x["month"])
        )

    return quarters


def get_half_yearly_label(month_num: int, year: int) -> str:
    if month_num in (4, 5, 6, 7, 8, 9):
        return f"H1 {year} (Apr–Sep)"
    else:
        fy_year = year if month_num in (10, 11, 12) else year - 1
        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"


def discover_half_yearly_files(data_dir: str) -> dict:
    monthly = discover_monthly_files(data_dir)
    half_years = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        hy_label = get_half_yearly_label(month_num, year)
        if hy_label not in half_years:
            half_years[hy_label] = []

        half_years[hy_label].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for hy in half_years:
        half_years[hy] = sorted(
            half_years[hy],
            key=lambda x: (x["year"], x["month"])
        )

    return half_years


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-MONTH VRM AND DRM COMBINE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_selected_vrm_data(selected_entries: list) -> dict:
    active_dfs, dropped_dfs, new_reg_dfs = [], [], []

    for entry in selected_entries:
        vrm_path = entry.get("vrm")
        if vrm_path and os.path.exists(vrm_path):
            month_data = load_data(vrm_path)
            if month_data:
                if "active" in month_data and not month_data["active"].empty:
                    df_act = month_data["active"].copy()
                    df_act["Reporting Period"] = entry["label"]
                    active_dfs.append(df_act)
                
                if "dropped" in month_data and not month_data["dropped"].empty:
                    df_drp = month_data["dropped"].copy()
                    df_drp["Reporting Period"] = entry["label"]
                    dropped_dfs.append(df_drp)
                
                if "new_reg" in month_data and not month_data["new_reg"].empty:
                    df_nr = month_data["new_reg"].copy()
                    df_nr["Reporting Period"] = entry["label"]
                    new_reg_dfs.append(df_nr)

    return {
        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),
        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),
        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),
    }


def load_selected_drm_data(selected_entries: list) -> dict:
    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []

    for entry in selected_entries:
        drm_path = entry.get("drm")
        if drm_path and os.path.exists(drm_path):
            m_drm = load_drm_data(drm_path)
            if m_drm:
                for key, container in [
                    ("sess", sess_dfs),
                    ("drm", drm_dfs),
                    ("ac", ac_dfs),
                    ("od", od_dfs),
                    ("ne", ne_dfs),
                ]:
                    if key in m_drm and not m_drm[key].empty:
                        df_sheet = m_drm[key].copy()
                        df_sheet["Reporting Period"] = entry["label"]
                        container.append(df_sheet)

    return {
        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),
        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),
        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),
        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),
        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),
    }


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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
# VRM DATA LOADING (SAFE PARSING)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names

    active_name = _find_sheet(xl, SHEET_ACTIVE)
    if active_name in sheets:
        active = pd.read_excel(xl, sheet_name=active_name)
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
    else:
        active = pd.DataFrame()

    dropped_name = _find_sheet(xl, SHEET_DROPPED)
    if dropped_name in sheets:
        dropped = pd.read_excel(xl, sheet_name=dropped_name)
        dropped.columns = dropped.columns.str.strip()
        if "State" in dropped.columns:
            dropped["State"] = (
                dropped["State"].astype(str)
                .str.replace("\u00ac\u00a0", " ", regex=False)
                .str.replace("\u00a0", " ", regex=False)
                .str.strip()
            )
    else:
        dropped = pd.DataFrame()

    new_reg_name = _find_sheet(xl, SHEET_NEW_REG)
    if new_reg_name in sheets:
        new_reg = pd.read_excel(xl, sheet_name=new_reg_name)
        new_reg.columns = new_reg.columns.str.strip()
        if "Date Joined" in new_reg.columns:
            new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
        if "Reference" in new_reg.columns:
            new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
        if "Gender" in new_reg.columns:
            new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()
    else:
        new_reg = pd.DataFrame()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

# ─────────────────────────────────────────────────────────────────────────────
# DRM DATA LOADING (SAFE PARSING)
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

DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

    sess_name = _find_sheet(xl, "SESSION DUMP")
    if sess_name in sheets:
        sess = pd.read_excel(xl, sheet_name=sess_name)
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
    else:
        sess = pd.DataFrame()

    drm_name = _find_sheet(xl, "DRM")
    if drm_name in sheets:
        drm = pd.read_excel(xl, sheet_name=drm_name)
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
    else:
        drm = pd.DataFrame()

    ac_name = _find_sheet(xl, "Active centers")
    ac = pd.read_excel(xl, sheet_name=ac_name) if ac_name in sheets else pd.DataFrame()
    if not ac.empty: ac.columns = ac.columns.str.strip()

    od_name = _find_sheet(xl, "Offering Details")
    od = pd.read_excel(xl, sheet_name=od_name) if od_name in sheets else pd.DataFrame()
    if not od.empty:
        od.columns = od.columns.str.strip()
        od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

    ne_name = _find_sheet(xl, "New Enrolled student")
    ne = pd.read_excel(xl, sheet_name=ne_name) if ne_name in sheets else pd.DataFrame()
    if not ne.empty: ne.columns = ne.columns.str.strip()

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

    # ─────────────────────────────────────────────────────────────────────────
    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)
    # ─────────────────────────────────────────────────────────────────────────

    reporting_type = st.radio(
        "📊 Reporting Type",
        ["Monthly", "Quarterly", "Half-Yearly"],
        horizontal=True,
        key="reporting_type",
    )

    selected_entries = []
    period_label = ""

    if reporting_type == "Monthly":
        month_labels = list(available_months.keys())

        selected_months = st.multiselect(
            "📅 Select Reporting Months",
            options=month_labels,
            default=[month_labels[-1]] if month_labels else [],
            key="selected_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            {
                "label": month,
                "vrm": available_months[month]["vrm"],
                "drm": available_months[month]["drm"],
            }
            for month in selected_months
        ]

        period_label = ", ".join(selected_months)

    elif reporting_type == "Quarterly":
        available_quarters = discover_quarterly_files(DATA_DIR)

        if not available_quarters:
            st.error("⚠️ No quarterly data available.")
            return

        quarter_labels = list(available_quarters.keys())

        selected_quarter = st.selectbox(
            "📅 Reporting Quarter",
            options=quarter_labels,
            index=len(quarter_labels) - 1,
            key="selected_quarter",
        )

        quarter_entries = available_quarters[selected_quarter]

        month_options = [
            item["label"]
            for item in quarter_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_quarter_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in quarter_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_quarter} — "
            + ", ".join(selected_months)
        )

    else:
        available_half_years = discover_half_yearly_files(DATA_DIR)

        if not available_half_years:
            st.error("⚠️ No half-yearly data available.")
            return

        hy_labels = list(available_half_years.keys())

        selected_hy = st.selectbox(
            "📅 Reporting Half-Year",
            options=hy_labels,
            index=len(hy_labels) - 1,
            key="selected_half_year",
        )

        hy_entries = available_half_years[selected_hy]

        month_options = [
            item["label"]
            for item in hy_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_hy_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in hy_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_hy} — "
            + ", ".join(selected_months)
        )

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {period_label}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD VRM & DRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading VRM & DRM data for {period_label}…"):
        data = load_selected_vrm_data(selected_entries)
        drm_data_raw = load_selected_drm_data(selected_entries)

    df_raw  = data.get("active", pd.DataFrame())
    dropped = data.get("dropped", pd.DataFrame())
    new_reg = data.get("new_reg", pd.DataFrame())

    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ── Build sidebar ─────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")

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

        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique()) if not df_raw.empty else []
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique()) if not df_raw.empty else []

            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

            if sel_donors and not df_raw.empty:
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

            _vmask = pd.Series([True] * len(df_raw), index=df_raw.index) if not df_raw.empty else []
            if sel_donors and not df_raw.empty:
                _vmask &= df_raw["Donor"].isin(sel_donors)
            if sel_states and not df_raw.empty:
                _vmask &= df_raw["State"].isin(sel_states)
            _linked_subjects = sorted(df_raw[_vmask]["Subject_clean"].dropna().unique()) if not df_raw.empty else []

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

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique()) if not df_raw.empty else []
            sel_states   = sorted(df_raw["State"].dropna().unique()) if not df_raw.empty else []
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique()) if not df_raw.empty else []

    if not df_raw.empty:
        df = df_raw[
            df_raw["Donor"].isin(sel_donors) &
            df_raw["State"].isin(sel_states) &
            df_raw["Subject_clean"].isin(sel_subjects)
        ].copy()
    else:
        df = pd.DataFrame()

    vol_df = df.drop_duplicates(subset="Volunteer id").copy() if not df.empty else df.copy()

    # ─────────────────────────────────────────────────────────────────────────
    # DECLARE TABS AT TOP-LEVEL FOR GUARANTEED RENDERING
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
            st.info("🔍 No VRM data available for the selected filters or month files.")
        else:
            total_vols_offering = len(df)
            unique_vols         = vol_df["Volunteer id"].nunique()
            dropped_vols        = len(dropped)
            new_vols            = new_reg["User ID"].nunique() if "User ID" in new_reg.columns else 0
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
            if "Date Joined" in new_reg.columns and "User ID" in new_reg.columns:
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
            if "Gender" in new_reg.columns and "User ID" in new_reg.columns:
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
            st.info("🔍 No VRM data available for the selected filters.")
        else:
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
                st.info("🔍 Gender breakdown (En Boys / En Girls) isn't available in this export.")

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
            st.info("🔍 No VRM data available for the selected filters.")
        else:
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
    # TAB 4 — DRM CLIENT REPORT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:
        if not drm_data_raw or "sess" not in drm_data_raw or drm_data_raw["sess"].empty:
            st.error(f"⚠️ No DRM file found or DRM data is empty for **{period_label}**.")
        else:
            sess_raw = drm_data_raw["sess"]
            drm_raw  = drm_data_raw["drm"]
            od_raw   = drm_data_raw["od"]

            sess = sess_raw[
                sess_raw["Donor"].isin(sel_drm_donors) &
                sess_raw["State"].isin(sel_drm_states) &
                sess_raw["Subject_clean"].isin(sel_drm_subjects)
            ].copy() if not sess_raw.empty else pd.DataFrame()

            drm = drm_raw[
                drm_raw["Donor Name"].isin(sel_drm_donors) &
                drm_raw["State"].isin(sel_drm_states)
            ].copy() if not drm_raw.empty else pd.DataFrame()

            if sess.empty or drm.empty:
                st.info("🔍 No DRM sessions available for the selected filters.")
            else:
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
                                    &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {period_label}
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

                total_sessions = len(sess)
                comp_sessions  = int((sess["Session_status"] == "Completed").sum())
                off_sessions   = int((sess["Session_status"] == "Offline").sum())
                canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
                total_clh_drm  = int(sess["Present/CLH"].sum())
                total_students = int(drm["Enrolled"].sum())
                avg_att_drm    = sess["Attendance%"].mean()
                total_centres  = drm["Center Name"].nunique()
                drm_total_vols = sess["Volunteer_id"].nunique() if "Volunteer_id" in sess.columns else 0
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

                for col_w, lbl, val, color, sub in [
                    (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
                    (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique active"),
                    (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned in period"),
                    (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
                    (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

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
                    (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{(off_sessions/total_sessions*100) if total_sessions else 0:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("---")

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

                st.markdown("#### ❌ Cancellations & CLH Split")

                col_ca1, col_ca2 = st.columns([1, 1.4])

                with col_ca1:
                    st.markdown("##### Cancellation reasons")
                    if "Cancel_reason" in sess.columns:
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
                    else:
                        st.info("Cancellation reasons are not available in this dataset.")

                with col_ca2:
                    st.markdown("##### State-wise CLH split")
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

                st.markdown("#### 📚 Subject & Grade Analytics")

                col_sg1, col_sg2 = st.columns(2)

                with col_sg1:
                    st.markdown("##### Sessions by subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
                        subj_grade_sess = (
                            sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                            .count().reset_index()
                            .rename(columns={"Session_id": "Sessions"})
                        )
                        pivot_sgs = (
                            subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                            .fillna(0)
                        )
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

                with col_sg2:
                    st.markdown("##### Session completion % — subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
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
                        if 'pivot_sgs' in locals():
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

                st.markdown("#### ⏱️ CLH by Subject & Grade")

                col_clh1, col_clh2 = st.columns([1.4, 1])

                with col_clh1:
                    st.markdown("##### CLH heatmap — subject × grade")
                    if "Grade" in sess.columns:
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

                st.markdown("#### 🙋 Volunteer Performance Scorecard")

                if {"Volunteer_id", "Teacher_name", "Session_id"}.issubset(sess.columns):
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

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                            padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                        eVidyaloka Programme Impact Report &nbsp;·&nbsp; {period_label}
                        &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                        {drm_total_vols} volunteers &nbsp;·&nbsp;
                        {total_students:,} students &nbsp;·&nbsp;
                        {total_clh_drm:,} child learning hours
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — TOPIC & SUBTOPIC ANALYTICS (INTEGRATED TOPIC DASHBOARD)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("### 📚 Topic & SubTopic Analytics Dashboard")

        uploaded_topic_file = st.file_uploader(
            "📁 Upload Topic/SubTopic Excel File",
            type=["xlsx", "xls"],
            key="topic_file_uploader",
            help="Upload the Excel file containing 'Topic-SubTopic', 'Cancelled Sessions', and 'Offline Sessions' sheets."
        )

        if uploaded_topic_file is None:
            st.info("👆 Please upload a Topic/SubTopic Excel file above to activate this dashboard.")
        else:
            @st.cache_data
            def load_topic_uploaded_data(file):
                try:
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names

                    def _get_sheet_name(target):
                        norm_target = " ".join(str(target).strip().lower().split())
                        for sheet in sheets:
                            if " ".join(str(sheet).strip().lower().split()) == norm_target:
                                return sheet
                        return target

                    t_sheet = _get_sheet_name("Topic-SubTopic")
                    c_sheet = _get_sheet_name("Cancelled Sessions")
                    o_sheet = _get_sheet_name("Offline Sessions")

                    t_df = pd.read_excel(file, sheet_name=t_sheet) if t_sheet in sheets else pd.DataFrame()
                    c_df = pd.read_excel(file, sheet_name=c_sheet) if c_sheet in sheets else pd.DataFrame()
                    o_df = pd.read_excel(file, sheet_name=o_sheet) if o_sheet in sheets else pd.DataFrame()

                    return t_df, c_df, o_df
                except Exception as e:
                    st.error(f"Error reading uploaded file: {e}")
                    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            topic_df, cancelled_df, offline_df = load_topic_uploaded_data(uploaded_topic_file)

            if topic_df.empty:
                st.error("The uploaded file does not contain a valid 'Topic-SubTopic' sheet or has no data.")
            else:
                # ── IN-TAB FILTERS ──────────────────────────────────────────
                with st.expander("📊 Topic Filters", expanded=True):
                    fc1, fc2, fc3 = st.columns(3)

                    state_col   = "state" if "state" in topic_df.columns else "State"
                    subject_col = "subject" if "subject" in topic_df.columns else "Subject"
                    grade_col   = "grade" if "grade" in topic_df.columns else "Grade"

                    with fc1:
                        st_opts = sorted(topic_df[state_col].dropna().unique()) if state_col in topic_df.columns else []
                        t_state_filter = st.multiselect("State", st_opts, key="tab5_state")

                    state_df = topic_df[topic_df[state_col].isin(t_state_filter)] if t_state_filter and state_col in topic_df.columns else topic_df.copy()

                    with fc2:
                        sub_opts = sorted(state_df[subject_col].dropna().unique()) if subject_col in state_df.columns else []
                        t_subject_filter = st.multiselect("Subject", sub_opts, key="tab5_subject")

                    subject_df = state_df[state_df[subject_col].isin(t_subject_filter)] if t_subject_filter and subject_col in state_df.columns else state_df.copy()

                    with fc3:
                        grd_opts = sorted(subject_df[grade_col].dropna().unique()) if grade_col in subject_df.columns else []
                        t_grade_filter = st.multiselect("Grade", grd_opts, key="tab5_grade")

                    filtered_topic_df = subject_df[subject_df[grade_col].isin(t_grade_filter)] if t_grade_filter and grade_col in subject_df.columns else subject_df.copy()

                # ── KPI METRICS ──────────────────────────────────────────────
                total_sessions   = len(filtered_topic_df)
                unique_topics    = filtered_topic_df["topic_name"].nunique() if "topic_name" in filtered_topic_df.columns else 0
                unique_subtopics = filtered_topic_df["sub_topic_name"].nunique() if "sub_topic_name" in filtered_topic_df.columns else 0
                unique_centers   = filtered_topic_df["center_id"].nunique() if "center_id" in filtered_topic_df.columns else (
                    filtered_topic_df["center_name"].nunique() if "center_name" in filtered_topic_df.columns else 0
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Sessions", f"{total_sessions:,}")
                with col2:
                    st.metric("Unique Topics", f"{unique_topics:,}")
                with col3:
                    st.metric("Unique Subtopics", f"{unique_subtopics:,}")
                with col4:
                    st.metric("Active Centers", f"{unique_centers:,}")

                st.markdown("---")

                # ── DASHBOARD SUB-TABS ────────────────────────────────────────
                top_tab1, top_tab2, top_tab3 = st.tabs([
                    "📋 Topic Analysis",
                    "❌ Cancelled Sessions",
                    "⏸️ Offline Sessions"
                ])

                # 1. TOPIC ANALYSIS
                with top_tab1:
                    if "topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Topics")
                        topic_summary = (
                            filtered_topic_df
                            .groupby("topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig = px.bar(
                            topic_summary, x="Sessions", y="topic_name",
                            orientation="h", title="Top 20 Topics"
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    if "sub_topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Sub Topics")
                        subtopic_summary = (
                            filtered_topic_df
                            .groupby("sub_topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig2 = px.bar(
                            subtopic_summary, x="Sessions", y="sub_topic_name",
                            orientation="h", title="Top 20 Sub Topics"
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    if subject_col in filtered_topic_df.columns:
                        st.subheader("Subject Distribution")
                        subject_chart = (
                            filtered_topic_df
                            .groupby(subject_col)
                            .size()
                            .reset_index(name="Count")
                        )
                        fig3 = px.pie(
                            subject_chart, names=subject_col, values="Count"
                        )
                        st.plotly_chart(fig3, use_container_width=True)

                # 2. CANCELLED SESSIONS
                with top_tab2:
                    st.subheader("Cancelled Session Analysis")
                    if not cancelled_df.empty:
                        st.metric("Total Cancelled Sessions", f"{len(cancelled_df):,}")

                        cancel_reason_col = (
                            "cancel_reason" if "cancel_reason" in cancelled_df.columns else (
                                "Cancel_reason" if "Cancel_reason" in cancelled_df.columns else None
                            )
                        )
                        if cancel_reason_col:
                            cancel_chart = (
                                cancelled_df
                                .groupby(cancel_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                                .head(15)
                            )
                            fig4 = px.bar(
                                cancel_chart, x="Count", y=cancel_reason_col,
                                orientation="h", title="Cancellation Reasons"
                            )
                            st.plotly_chart(fig4, use_container_width=True)

                        canc_state_col = "State" if "State" in cancelled_df.columns else (
                            "state" if "state" in cancelled_df.columns else None
                        )
                        if canc_state_col:
                            state_cancel = (
                                cancelled_df
                                .groupby(canc_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig5 = px.bar(
                                state_cancel, x=canc_state_col, y="Count",
                                title="Cancelled Sessions by State"
                            )
                            st.plotly_chart(fig5, use_container_width=True)
                    else:
                        st.info("No cancelled sessions data found in the uploaded file.")

                # 3. OFFLINE SESSIONS
                with top_tab3:
                    st.subheader("Offline Session Analysis")
                    if not offline_df.empty:
                        st.metric("Total Offline Sessions", f"{len(offline_df):,}")

                        offline_reason_col = (
                            "Offline reason" if "Offline reason" in offline_df.columns else (
                                "offline_reason" if "offline_reason" in offline_df.columns else None
                            )
                        )
                        if offline_reason_col:
                            offline_reason = (
                                offline_df
                                .groupby(offline_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig6 = px.bar(
                                offline_reason, x="Count", y=offline_reason_col,
                                orientation="h", title="Offline Reasons"
                            )
                            st.plotly_chart(fig6, use_container_width=True)

                        off_state_col = "State" if "State" in offline_df.columns else (
                            "state" if "state" in offline_df.columns else None
                        )
                        if off_state_col:
                            state_offline = (
                                offline_df
                                .groupby(off_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig7 = px.bar(
                                state_offline, x=off_state_col, y="Count",
                                title="Offline Sessions by State"
                            )
                            st.plotly_chart(fig7, use_container_width=True)
                    else:
                        st.info("No offline sessions data found in the uploaded file.")

                # ── EXPORT DATA ──────────────────────────────────────────────
                st.markdown("---")
                st.download_button(
                    "📥 Download Filtered Topic Data",
                    filtered_topic_df.to_csv(index=False),
                    file_name=f"filtered_{uploaded_topic_file.name.replace('.xlsx', '.csv')}",
                    mime="text/csv",
                    key="download_topic_csv"
                )
in my ops dashboard file in impact analytics dashboard i want to see the baseline assesment visuals is it possible if have only BL data

Yes, it is fully possible to display baseline assessment visuals even if you only have Baseline (BL) data.

While you won't be able to calculate Endline vs Baseline growth/delta, having single-point baseline data allows you to extract rich diagnostic insights about student preparedness before academic interventions begin.

What You Can Visualize With Baseline Data Only
Student Baseline Score Distribution:

Score range breakdowns (e.g., <40% Needs Support, 40–70% Developing, >70% Proficient).

Box plots or histograms showing score distributions per subject or grade.

Subject & Grade Preparedness:

Average baseline percentage score per Subject (Math, Science, English).

Average score comparison across Grades (Grade 5 to 10).

Geographic & Centre Level Diagnostics:

State-wise average baseline performance heatmap.

Lowest performing centres that require immediate academic focus.

Gender & Demographic Breakdown:

Comparison of baseline performance between Boys vs Girls (if gender data is available in the sheet).

Implementation Options for ops_dashboard.py
You can integrate this in your app in two ways:

Option 1: File Uploader inside Tab 5 or a dedicated "Baseline Analytics" Tab
Allow uploading the standalone Baseline Assessment file (.xlsx) via st.file_uploader() (similar to how Tab 5 works).

Option 2: Add auto-detection logic in discover_monthly_files
Add matching logic for BL_<Month>_<Year>.xlsx files to load baseline sheets directly alongside VRM and DRM files.

Python Visual Snippet (Streamlit + Plotly)
Here is a ready-to-use Python snippet showing how to build standard baseline visualizations in Streamlit when given a DataFrame bl_df:

Python
import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Define Proficiency Buckets
def get_bucket(score):
    if score < 40: return "🔴 Needs Support (<40%)"
    elif score < 70: return "🟡 Developing (40-70%)"
    else: return "🟢 Proficient (>70%)"

bl_df["Performance_Category"] = bl_df["Baseline%"].apply(get_bucket)

# 2. Key Metrics Top Bar
col1, col2, col3, col4 = st.columns(4)
col1.metric("Students Assessed", f"{bl_df['Student_ID'].nunique():,}")
col2.metric("Avg Baseline Score", f"{bl_df['Baseline%'].mean():.1f}%")
col3.metric("Pass / Proficient Rate", f"{(bl_df['Baseline%'] >= 40).mean()*100:.1f}%")
col4.metric("Centres Covered", f"{bl_df['Center'].nunique():,}")

st.markdown("---")

# 3. Chart Columns
c1, c2 = st.columns(2)

with c1:
    st.subheader("📊 Performance Tier Distribution")
    cat_counts = bl_df["Performance_Category"].value_counts().reset_index()
    cat_counts.columns = ["Tier", "Students"]
    fig_pie = px.pie(
        cat_counts, names="Tier", values="Students", hole=0.4,
        color="Tier",
        color_discrete_map={
            "🔴 Needs Support (<40%)": "#ed1c2d",
            "🟡 Developing (40-70%)": "#fdcb6e",
            "🟢 Proficient (>70%)": "#00964d"
        }
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    st.subheader("📚 Avg Baseline Score by Subject & Grade")
    sub_grade = bl_df.groupby(["Subject", "Grade"])["Baseline%"].mean().reset_index()
    fig_bar = px.bar(
        sub_grade, x="Grade", y="Baseline%", color="Subject",
        barmode="group", text_auto=".1f"
    )
    fig_bar.update_layout(yaxis_range=[0, 100])
    st.plotly_chart(fig_bar, use_container_width=True)
Would you like help creating a dedicated Baseline Assessment Tab with a specific Excel structure, or expanding the file auto-loader to parse Baseline assessment sheets alongside VRM/DRM?

i have already working code BL N EL but i have now only BL data so can i just upload it and see visuals

Yes, you can upload only the Baseline (BL) file to see the visuals, but the dashboard will throw an error if any of your visualization functions attempt to compute comparison metrics like:

Python
# ❌ This will crash if EL data is missing:
growth = el_df["Score"] - bl_df["Score"]
To make your existing code work seamlessly with BL data only, you need to wrap the Endline/Growth sections in simple checks so the app gracefully falls back to Baseline Diagnostic Mode.

Step 1: Add a Data-Availability Safety Check
In ops_dashboard.py, inspect the data loader where you load the Baseline and Endline sheets. Update the logic to detect whether Endline data is present:

Python
# Check if Endline data actually exists and has rows
has_endline = (
    "endline" in data 
    and isinstance(data["endline"], pd.DataFrame) 
    and not data["endline"].empty
)
Step 2: Guard Endline & Comparison Visuals
Wrap any comparative KPIs, growth charts, or Endline-specific figures with if has_endline: checks.

Here is how to structure your tab code so it displays Baseline visuals when only BL data is uploaded:

Python
# 1. BASELINE VISUALS (Always Rendered)
st.markdown("#### 📊 Baseline Diagnostics")

col1, col2, col3 = st.columns(3)
col1.metric("Students Assessed (BL)", f"{len(bl_df):,}")
col2.metric("Avg Baseline Score", f"{bl_df['Baseline%'].mean():.1f}%")
col3.metric("Centres Covered", f"{bl_df['Center'].nunique():,}")

# Baseline Score Distribution Chart
fig_bl = px.histogram(
    bl_df, x="Baseline%", nbins=20, 
    title="Baseline Score Distribution",
    color_discrete_sequence=["#185fa5"]
)
st.plotly_chart(fig_bl, use_container_width=True)


# 2. COMPARATIVE / ENDLINE VISUALS (Rendered ONLY if EL is present)
if has_endline:
    st.markdown("#### 📈 Impact & Growth (BL vs EL)")
    
    el_df = data["endline"]
    avg_growth = el_df["Endline%"].mean() - bl_df["Baseline%"].mean()
    st.metric("Avg Score Growth", f"+{avg_growth:.1f}%")
    
    # Render comparative charts here...
else:
    st.info("ℹ️ **Baseline-Only Mode Active:** Upload Endline data when available to unlock BL vs EL growth & impact analytics.")
Step 3: Handle Missing Column References
If your existing filters or charts reference columns that only exist in the Endline sheet (such as Endline%, Delta, or Growth_Category), use safe fallback gets when filtering:

Python
# Example of safe column selection:
score_col = "Endline%" if (has_endline and "Endline%" in df.columns) else "Baseline%"
Once you add the has_endline check around your comparison code, you can upload your BL file directly and view all Baseline distribution, subject, grade, and centre-level diagnostics without any errors.

in my ops dashboard there is baseline nd endline comparison tab is there but i also want baseline only metrics tab can add that patch of code .below is the code . """

ops_dashboard.py

================

Operations & Impact Command Center  —  eVidyaloka VRM

Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /

DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,

Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().



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

DATA_DIR = os.path.dirname(os.path.abspath(__file__))



# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.

MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)



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

    them by reporting month.

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





# ─────────────────────────────────────────────────────────────────────────────

# QUARTERLY & HALF-YEARLY REPORTING HELPERS

# ─────────────────────────────────────────────────────────────────────────────



def get_quarter_label(month_num: int, year: int) -> str:

    """Financial year quarters:

       Q1 = Apr-Jun

       Q2 = Jul-Sep

       Q3 = Oct-Dec

       Q4 = Jan-Mar

    """

    if month_num in (4, 5, 6):

        return f"Q1 {year} (Apr–Jun)"

    elif month_num in (7, 8, 9):

        return f"Q2 {year} (Jul–Sep)"

    elif month_num in (10, 11, 12):

        return f"Q3 {year} (Oct–Dec)"

    else:

        return f"Q4 {year - 1}-{year} (Jan–Mar)"





def discover_quarterly_files(data_dir: str) -> dict:

    """Group available monthly VRM/DRM files into financial quarters."""

    monthly = discover_monthly_files(data_dir)

    quarters = {}



    for label, paths in monthly.items():

        month_name, year = label.rsplit(" ", 1)

        month_num = _month_sort_key(month_name)

        year = int(year)



        if month_num == 0:

            continue



        quarter = get_quarter_label(month_num, year)

        if quarter not in quarters:

            quarters[quarter] = []



        quarters[quarter].append({

            "label": label,

            "month": month_num,

            "year": year,

            "vrm": paths["vrm"],

            "drm": paths["drm"],

        })



    for quarter in quarters:

        quarters[quarter] = sorted(

            quarters[quarter],

            key=lambda x: (x["year"], x["month"])

        )



    return quarters





def get_half_yearly_label(month_num: int, year: int) -> str:

    """Financial year half-years:

       H1 = Apr-Sep

       H2 = Oct-Mar

    """

    if month_num in (4, 5, 6, 7, 8, 9):

        return f"H1 {year} (Apr–Sep)"

    else:

        fy_year = year if month_num in (10, 11, 12) else year - 1

        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"





def discover_half_yearly_files(data_dir: str) -> dict:

    """Group available monthly VRM/DRM files into financial half-years."""

    monthly = discover_monthly_files(data_dir)

    half_years = {}



    for label, paths in monthly.items():

        month_name, year = label.rsplit(" ", 1)

        month_num = _month_sort_key(month_name)

        year = int(year)



        if month_num == 0:

            continue



        hy_label = get_half_yearly_label(month_num, year)

        if hy_label not in half_years:

            half_years[hy_label] = []



        half_years[hy_label].append({

            "label": label,

            "month": month_num,

            "year": year,

            "vrm": paths["vrm"],

            "drm": paths["drm"],

        })



    for hy in half_years:

        half_years[hy] = sorted(

            half_years[hy],

            key=lambda x: (x["year"], x["month"])

        )



    return half_years





# ─────────────────────────────────────────────────────────────────────────────

# MULTI-MONTH VRM AND DRM COMBINE LOADERS

# ─────────────────────────────────────────────────────────────────────────────



def load_selected_vrm_data(selected_entries: list) -> dict:

    """

    Loads VRM data for multiple selected month entries and concatenates 

    them together across 'active', 'dropped', and 'new_reg' sheets.

    """

    active_dfs, dropped_dfs, new_reg_dfs = [], [], []



    for entry in selected_entries:

        vrm_path = entry.get("vrm")

        if vrm_path and os.path.exists(vrm_path):

            month_data = load_data(vrm_path)

            if month_data:

                if "active" in month_data and not month_data["active"].empty:

                    df_act = month_data["active"].copy()

                    df_act["Reporting Period"] = entry["label"]

                    active_dfs.append(df_act)

                

                if "dropped" in month_data and not month_data["dropped"].empty:

                    df_drp = month_data["dropped"].copy()

                    df_drp["Reporting Period"] = entry["label"]

                    dropped_dfs.append(df_drp)

                

                if "new_reg" in month_data and not month_data["new_reg"].empty:

                    df_nr = month_data["new_reg"].copy()

                    df_nr["Reporting Period"] = entry["label"]

                    new_reg_dfs.append(df_nr)



    return {

        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),

        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),

        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),

    }





def load_selected_drm_data(selected_entries: list) -> dict:

    """

    Loads DRM data for multiple selected month entries and concatenates 

    them across all individual DRM sheets.

    """

    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []



    for entry in selected_entries:

        drm_path = entry.get("drm")

        if drm_path and os.path.exists(drm_path):

            m_drm = load_drm_data(drm_path)

            if m_drm:

                for key, container in [

                    ("sess", sess_dfs),

                    ("drm", drm_dfs),

                    ("ac", ac_dfs),

                    ("od", od_dfs),

                    ("ne", ne_dfs),

                ]:

                    if key in m_drm and not m_drm[key].empty:

                        df_sheet = m_drm[key].copy()

                        df_sheet["Reporting Period"] = entry["label"]

                        container.append(df_sheet)



    return {

        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),

        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),

        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),

        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),

        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),

    }





def _find_sheet(xl: pd.ExcelFile, target: str) -> str:

    def _norm(s: str) -> str:

        return " ".join(str(s).strip().lower().split())



    target_norm = _norm(target)

    for name in xl.sheet_names:

        if _norm(name) == target_norm:

            return name

    return target





def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:

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

    if not os.path.exists(path):

        return {}



    xl = pd.ExcelFile(path)



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



    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))

    dropped.columns = dropped.columns.str.strip()

    if "State" in dropped.columns:

        dropped["State"] = (

            dropped["State"].astype(str)

            .str.replace("\u00ac\u00a0", " ", regex=False)

            .str.replace("\u00a0", " ", regex=False)

            .str.strip()

        )



    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))

    new_reg.columns = new_reg.columns.str.strip()

    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")

    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)

    if "Gender" in new_reg.columns:

        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()



    return {"active": active, "dropped": dropped, "new_reg": new_reg}



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



DRM_SHEET_ALIASES = {

    "Donor Name":     ["Donor"],

    "Live_CLH":       ["Live CLH"],

    "Attendance %":   ["Live Attendance %"],

}





@st.cache_data(show_spinner=False)

def load_drm_data(path: str) -> dict:

    if not os.path.exists(path):

        return {}



    xl = pd.ExcelFile(path)



    def _fix_state(s):

        return (str(s)

                .replace("\u00ac\u00a0", " ")

                .replace("\u00a0", " ")

                .strip())



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



    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))

    ac.columns = ac.columns.str.strip()



    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))

    od.columns = od.columns.str.strip()

    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])



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



    # ─────────────────────────────────────────────────────────────────────────

    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)

    # ─────────────────────────────────────────────────────────────────────────



    reporting_type = st.radio(

        "📊 Reporting Type",

        ["Monthly", "Quarterly", "Half-Yearly"],

        horizontal=True,

        key="reporting_type",

    )



    selected_entries = []

    period_label = ""



    if reporting_type == "Monthly":

        month_labels = list(available_months.keys())



        selected_months = st.multiselect(

            "📅 Select Reporting Months",

            options=month_labels,

            default=[month_labels[-1]] if month_labels else [],

            key="selected_months",

            placeholder="Select one or more months",

        )



        if not selected_months:

            st.warning("⚠️ Please select at least one month.")

            return



        selected_entries = [

            {

                "label": month,

                "vrm": available_months[month]["vrm"],

                "drm": available_months[month]["drm"],

            }

            for month in selected_months

        ]



        period_label = ", ".join(selected_months)



    elif reporting_type == "Quarterly":

        available_quarters = discover_quarterly_files(DATA_DIR)



        if not available_quarters:

            st.error("⚠️ No quarterly data available.")

            return



        quarter_labels = list(available_quarters.keys())



        selected_quarter = st.selectbox(

            "📅 Reporting Quarter",

            options=quarter_labels,

            index=len(quarter_labels) - 1,

            key="selected_quarter",

        )



        quarter_entries = available_quarters[selected_quarter]



        month_options = [

            item["label"]

            for item in quarter_entries

        ]



        selected_months = st.multiselect(

            "📅 Select Months",

            options=month_options,

            default=month_options,

            key="selected_quarter_months",

            placeholder="Select one or more months",

        )



        if not selected_months:

            st.warning("⚠️ Please select at least one month.")

            return



        selected_entries = [

            item

            for item in quarter_entries

            if item["label"] in selected_months

        ]



        period_label = (

            f"{selected_quarter} — "

            + ", ".join(selected_months)

        )



    else:

        available_half_years = discover_half_yearly_files(DATA_DIR)



        if not available_half_years:

            st.error("⚠️ No half-yearly data available.")

            return



        hy_labels = list(available_half_years.keys())



        selected_hy = st.selectbox(

            "📅 Reporting Half-Year",

            options=hy_labels,

            index=len(hy_labels) - 1,

            key="selected_half_year",

        )



        hy_entries = available_half_years[selected_hy]



        month_options = [

            item["label"]

            for item in hy_entries

        ]



        selected_months = st.multiselect(

            "📅 Select Months",

            options=month_options,

            default=month_options,

            key="selected_hy_months",

            placeholder="Select one or more months",

        )



        if not selected_months:

            st.warning("⚠️ Please select at least one month.")

            return



        selected_entries = [

            item

            for item in hy_entries

            if item["label"] in selected_months

        ]



        period_label = (

            f"{selected_hy} — "

            + ", ".join(selected_months)

        )



    st.markdown(

        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"

        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"

        f"  —  {period_label}</p>",

        unsafe_allow_html=True,

    )

    st.markdown("---")



    # ═════════════════════════════════════════════════════════════════════════

    # LOAD VRM DATA

    # ═════════════════════════════════════════════════════════════════════════



    with st.spinner(f"Loading VRM data for {period_label}…"):

        data = load_selected_vrm_data(selected_entries)



    if not data or data["active"].empty:

        st.error(f"⚠️ Could not load VRM data for {period_label}.")

        return



    df_raw  = data["active"]

    dropped = data["dropped"]

    new_reg = data["new_reg"]



    if "active_tab" not in st.session_state:

        st.session_state.active_tab = "vrm"



    # ═════════════════════════════════════════════════════════════════════════

    # LOAD DRM DATA

    # ═════════════════════════════════════════════════════════════════════════



    with st.spinner(f"Loading DRM data for {period_label}…"):

        drm_data_raw = load_selected_drm_data(selected_entries)



    # ── Build sidebar — content switches based on active_tab flag ─────────────

    with st.sidebar:

        st.markdown("---")



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



        if st.session_state.active_tab == "vrm":

            st.markdown(

                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"

                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",

                unsafe_allow_html=True,

            )



            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())

            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())

            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())



            sel_donors = st.multiselect(

                "Donor",

                options=_vrm_all_donors,

                default=_vrm_all_donors,

                placeholder="Type to search donors…",

                key="ops_donor",

            )



            if sel_donors

Notice that your pasted code got truncated partway through render_ops_dashboard() (right after if sel_donors), so I can't see the exact column names of your existing Baseline/Endline comparison tab.

Here is a self-contained Baseline-Only Metrics Tab patch designed to drop directly into your ops_dashboard.py alongside your existing tabs. It auto-normalizes common column names (like Baseline%, Score%, Student_ID, Center, State, Subject, Grade, Gender) so it works regardless of slight header variations.

1. Update the st.tabs(...) Declaration
In render_ops_dashboard(), add "🎯 Baseline Only" to your existing st.tabs([...]) list (and unpack an extra tab variable, e.g. tab_bl):

Python
tab_vol, tab_ctr, tab_aca, tab_drm, tab_top, tab_bl = st.tabs([
    "🙋 Volunteers",
    "🏫 Centres",
    "📚 Academic Health",
    "📊 DRM Client Report",
    "📚 Topic & SubTopic Analytics",
    "🎯 Baseline Only",
])
2. Add the with tab_bl: Block
Paste this block at the bottom of render_ops_dashboard() right after your with tab_top: (or BL/EL comparison tab) block:

Python
    # ═════════════════════════════════════════════════════════════════════════
    # TAB — BASELINE ONLY METRICS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_bl:
        st.markdown("### 🎯 Baseline Assessment Diagnostics (BL Only)")
        st.caption("Upload a standalone Baseline assessment Excel file to inspect student readiness, proficiency tiers, and subject/grade gaps.")

        bl_file = st.file_uploader(
            "📁 Upload Baseline Assessment File (.xlsx)",
            type=["xlsx", "xls"],
            key="bl_only_uploader",
        )

        if bl_file is None:
            st.info("👆 Upload your Baseline Excel file above to view Baseline-only metrics and visuals.")
        else:
            @st.cache_data(show_spinner=False)
            def _load_bl_only(file):
                xl = pd.ExcelFile(file)
                # Prefer a sheet with 'baseline' or 'bl' in the name, else first sheet
                target_sheet = xl.sheet_names[0]
                for s in xl.sheet_names:
                    if any(k in s.lower() for k in ("baseline", "bl", "score", "data")):
                        target_sheet = s
                        break
                df_b = pd.read_excel(xl, sheet_name=target_sheet)
                df_b.columns = df_b.columns.str.strip()
                return df_b

            bl_raw = _load_bl_only(bl_file)

            if bl_raw.empty:
                st.error("⚠️ The uploaded Baseline sheet is empty.")
            else:
                # ── Flexible column resolver ─────────────────────────────────
                def _pick_col(df_check, candidates):
                    lower_map = {c.lower().strip(): c for c in df_check.columns}
                    for cand in candidates:
                        if cand.lower() in lower_map:
                            return lower_map[cand.lower()]
                    return None

                col_score   = _pick_col(bl_raw, ["Baseline%", "Baseline %", "BL%", "BL %", "Baseline Score", "Score%", "Score %", "Percentage", "Marks %"])
                col_student = _pick_col(bl_raw, ["Student ID", "Student_ID", "Student Name", "Child ID", "Enrollment ID"])
                col_center  = _pick_col(bl_raw, ["Center Name", "Center", "Centre Name", "Centre"])
                col_state   = _pick_col(bl_raw, ["State", "Center State"])
                col_subject = _pick_col(bl_raw, ["Subject", "Subject Name"])
                col_grade   = _pick_col(bl_raw, ["Grade", "Class"])
                col_gender  = _pick_col(bl_raw, ["Gender", "Sex"])

                if col_score is None:
                    st.error(
                        f"⚠️ Could not detect a percentage score column. Found columns: `{list(bl_raw.columns)}`. "
                        "Ensure a column like `Baseline%`, `BL %`, or `Percentage` exists."
                    )
                else:
                    bl_df = bl_raw.copy()
                    bl_df["_score"] = pd.to_numeric(bl_df[col_score], errors="coerce")
                    bl_df = bl_df.dropna(subset=["_score"])

                    # ── In-Tab Cascading Filters ─────────────────────────────
                    with st.expander("🔍 Baseline Filters", expanded=True):
                        f1, f2, f3 = st.columns(3)
                        with f1:
                            if col_state:
                                st_opts = sorted(bl_df[col_state].dropna().astype(str).unique())
                                sel_bl_state = st.multiselect("State", st_opts, default=st_opts, key="bl_only_state")
                                if sel_bl_state:
                                    bl_df = bl_df[bl_df[col_state].astype(str).isin(sel_bl_state)]
                        with f2:
                            if col_subject:
                                sub_opts = sorted(bl_df[col_subject].dropna().astype(str).unique())
                                sel_bl_sub = st.multiselect("Subject", sub_opts, default=sub_opts, key="bl_only_sub")
                                if sel_bl_sub:
                                    bl_df = bl_df[bl_df[col_subject].astype(str).isin(sel_bl_sub)]
                        with f3:
                            if col_grade:
                                grd_opts = sorted(bl_df[col_grade].dropna().astype(str).unique())
                                sel_bl_grd = st.multiselect("Grade", grd_opts, default=grd_opts, key="bl_only_grd")
                                if sel_bl_grd:
                                    bl_df = bl_df[bl_df[col_grade].astype(str).isin(sel_bl_grd)]

                    if bl_df.empty:
                        st.info("No records match the selected Baseline filters.")
                    else:
                        # ── Proficiency Buckets ──────────────────────────────
                        def _tier(v):
                            if v < 40:
                                return "Needs Support (<40%)"
                            elif v < 70:
                                return "Developing (40–69%)"
                            return "Proficient (≥70%)"

                        bl_df["Proficiency Tier"] = bl_df["_score"].apply(_tier)

                        # ── KPI Strip ────────────────────────────────────────
                        n_assessed  = bl_df[col_student].nunique() if col_student else len(bl_df)
                        avg_bl      = bl_df["_score"].mean()
                        med_bl      = bl_df["_score"].median()
                        prof_pct    = (bl_df["_score"] >= 70).mean() * 100
                        support_pct = (bl_df["_score"] < 40).mean() * 100
                        n_centres   = bl_df[col_center].nunique() if col_center else 0

                        k1, k2, k3, k4, k5 = st.columns(5)
                        k1.markdown(_kpi("Students Assessed", f"{n_assessed:,}",   P["teal"],   f"{n_centres} centres" if col_center else "total rows"), unsafe_allow_html=True)
                        k2.markdown(_kpi("Avg BL Score",      f"{avg_bl:.1f}%",     P["violet"], f"median {med_bl:.1f}%"), unsafe_allow_html=True)
                        k3.markdown(_kpi("Proficient (≥70%)", f"{prof_pct:.1f}%",   P["green"],  "ready / strong"), unsafe_allow_html=True)
                        k4.markdown(_kpi("Needs Support",     f"{support_pct:.1f}%",P["red"],    "below 40%"), unsafe_allow_html=True)
                        k5.markdown(_kpi("Total Assessments", f"{len(bl_df):,}",    P["amber"],  "subject-student tests"), unsafe_allow_html=True)

                        st.markdown("---")

                        # ── Row 1: Tier Donut + Score Histogram ──────────────
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("##### Proficiency Tier Breakdown")
                            tier_df = bl_df["Proficiency Tier"].value_counts().reset_index()
                            tier_df.columns = ["Tier", "Count"]
                            fig_tier = px.pie(
                                tier_df, names="Tier", values="Count", hole=0.48,
                                color="Tier",
                                color_discrete_map={
                                    "Needs Support (<40%)": P["red"],
                                    "Developing (40–69%)":  P["amber"],
                                    "Proficient (≥70%)":    P["green"],
                                },
                            )
                            fig_tier.update_traces(textposition="outside", texttemplate="<b>%{label}</b><br>%{percent:.1%}")
                            fig_tier.update_layout(height=320, showlegend=False, plot_bgcolor=BG, paper_bgcolor=BG, margin=dict(l=10, r=10, t=30, b=10))
                            st.plotly_chart(fig_tier, use_container_width=True)

                        with c2:
                            st.markdown("##### Baseline Score Distribution")
                            fig_bl_hist = px.histogram(
                                bl_df, x="_score", nbins=20,
                                color_discrete_sequence=[P["teal"]], opacity=0.85,
                            )
                            fig_bl_hist.add_vline(x=avg_bl, line_dash="dash", line_color=P["red"], annotation_text=f"Mean {avg_bl:.1f}%")
                            _layout(fig_bl_hist, height=320)
                            fig_bl_hist.update_layout(xaxis_title="Baseline Score %", yaxis_title="Assessments", bargap=0.05)
                            st.plotly_chart(fig_bl_hist, use_container_width=True)

                        st.markdown("---")

                        # ── Row 2: Subject & Grade Performance ───────────────
                        c3, c4 = st.columns(2)
                        with c3:
                            if col_subject:
                                st.markdown("##### Avg Baseline % by Subject")
                                sub_bl = (
                                    bl_df.groupby(col_subject)["_score"].mean()
                                    .reset_index().rename(columns={"_score": "Avg BL%"})
                                    .sort_values("Avg BL%", ascending=True)
                                )
                                fig_sub_bl = px.bar(
                                    sub_bl, x="Avg BL%", y=col_subject, orientation="h",
                                    text=sub_bl["Avg BL%"].round(1).astype(str) + "%",
                                    color="Avg BL%", color_continuous_scale=[P["coral"], P["amber"], P["green"]],
                                    range_color=[0, 100],
                                )
                                fig_sub_bl.update_traces(textposition="outside", marker_line_width=0)
                                fig_sub_bl.update_coloraxes(showscale=False)
                                _layout(fig_sub_bl, height=320, margin=dict(l=0, r=60, t=20, b=0))
                                fig_sub_bl.update_layout(xaxis=dict(range=[0, 112], title="Avg Baseline %"), yaxis_title="")
                                st.plotly_chart(fig_sub_bl, use_container_width=True)

                        with c4:
                            if col_subject and col_grade:
                                st.markdown("##### Subject × Grade Avg Baseline %")
                                sg_piv = bl_df.pivot_table(index=col_subject, columns=col_grade, values="_score", aggfunc="mean").fillna(0)
                                fig_sg_bl = px.imshow(
                                    sg_piv, text_auto=".0f", aspect="auto",
                                    color_continuous_scale=["#fadbd8", P["amber"], P["teal"]],
                                    zmin=0, zmax=100,
                                )
                                fig_sg_bl.update_layout(height=320, plot_bgcolor=BG, paper_bgcolor=BG, margin=dict(l=0, r=0, t=20, b=0), coloraxis_showscale=False)
                                st.plotly_chart(fig_sg_bl, use_container_width=True)

                        # ── Row 3: State & Gender Diagnostics ────────────────
                        if col_state or col_gender:
                            st.markdown("---")
                            c5, c6 = st.columns([1.4, 1])
                            with c5:
                                if col_state:
                                    st.markdown("##### State-wise Avg Baseline %")
                                    st_bl = (
                                        bl_df.groupby(col_state)["_score"].mean()
                                        .reset_index().rename(columns={"_score": "Avg BL%"})
                                        .sort_values("Avg BL%", ascending=True)
                                    )
                                    fig_st_bl = px.bar(
                                        st_bl, x="Avg BL%", y=col_state, orientation="h",
                                        text=st_bl["Avg BL%"].round(1).astype(str) + "%",
                                        color_discrete_sequence=[P["teal"]],
                                    )
                                    fig_st_bl.update_traces(textposition="outside", marker_line_width=0)
                                    _layout(fig_st_bl, height=max(260, len(st_bl) * 36), margin=dict(l=0, r=60, t=20, b=0))
                                    fig_st_bl.update_layout(xaxis=dict(range=[0, 112], title="Avg Baseline %"), yaxis_title="")
                                    st.plotly_chart(fig_st_bl, use_container_width=True)
                            with c6:
                                if col_gender and col_subject:
                                    st.markdown("##### Gender Comparison by Subject")
                                    gen_bl = bl_df.groupby([col_subject, col_gender])["_score"].mean().reset_index()
                                    fig_gen_bl = px.bar(
                                        gen_bl, x=col_subject, y="_score", color=col_gender,
                                        barmode="group", text_auto=".1f",
                                        color_discrete_map={"Male": P["teal"], "Female": P["coral"], "Boys": P["teal"], "Girls": P["coral"]},
                                    )
                                    _layout(fig_gen_bl, height=300, legend_bottom=True)
                                    fig_gen_bl.update_layout(yaxis=dict(range=[0, 105], title="Avg BL %"), xaxis_title="")
                                    st.plotly_chart(fig_gen_bl, use_container_width=True)

                        # ── Priority Centres Table (Lowest Avg BL) ───────────
                        if col_center:
                            st.markdown("---")
                            st.markdown("##### 🏫 Centre Baseline Leaderboard (Lowest to Highest — Priority Support)")
                            ctr_grp_cols = [col_center] + ([col_state] if col_state else [])
                            ctr_bl = (
                                bl_df.groupby(ctr_grp_cols).agg(
                                    Assessments=("_score", "count"),
                                    Avg_BL_Pct=("_score", "mean"),
                                    Needs_Support_Pct=("_score", lambda s: (s < 40).mean() * 100),
                                    Proficient_Pct=("_score", lambda s: (s >= 70).mean() * 100),
                                ).reset_index()
                                .sort_values("Avg_BL_Pct", ascending=True)
                            )
                            st.dataframe(
                                ctr_bl,
                                use_container_width=True,
                                hide_index=True,
                                height=380,
                                column_config={
                                    "Avg_BL_Pct": st.column_config.ProgressColumn("Avg BL %", min_value=0, max_value=100, format="%.1f%%"),
                                    "Needs_Support_Pct": st.column_config.NumberColumn("Needs Support %", format="%.1f%%"),
                                    "Proficient_Pct": st.column_config.NumberColumn("Proficient %", format="%.1f%%"),
                                },
                            )
How to Apply It
Locate your st.tabs([...]) line in ops_dashboard.py and append "🎯 Baseline Only" along with tab_bl on the left-hand side of =.

Paste the with tab_bl: section at the end of render_ops_dashboard() matching the 4-space indentation of your other with tab_...: blocks.

If your existing BL/EL comparison tab already has an uploader or loader for Baseline data and you want both tabs to share the same single file upload instead of uploading twice, share the bottom half of your ops_dashboard.py (from st.tabs(...) to the end of the file) and I will wire them together.

I have added Tab 6 (🎯 Baseline Only) to your ops_dashboard (15).py file.

Key Updates Made:
Added tab_bl ("🎯 Baseline Only") to st.tabs([...]) right alongside your existing 5 tabs.

Removed internal st.stop() calls inside tabs (replacing them with if/else blocks) so that if a filter returns empty in Tab 1–4, Streamlit does not halt execution before rendering Tab 5 and Tab 6.

Built the Baseline-Only Diagnostics Tab (with tab_bl:) featuring:

Automatic column detection (handles Baseline%, BL %, Score%, Student ID, Center Name, State, Subject, Grade, Gender).

Cascading filters for State, Subject, and Grade.

5 KPI cards: Students Assessed, Avg BL Score (with median), Proficient Rate (≥70%), Needs Support Rate (<40%), and Total Assessments.

Visuals: Proficiency Tier donut chart, Score Distribution histogram with mean marker, Subject horizontal bar chart, Subject × Grade heatmap, State-wise performance, Gender comparison, and a Priority Support Centre leaderboard.

Complete Updated ops_dashboard.py
Python
"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,
Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().

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
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

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
    them by reporting month.
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


# ─────────────────────────────────────────────────────────────────────────────
# QUARTERLY & HALF-YEARLY REPORTING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_quarter_label(month_num: int, year: int) -> str:
    """Financial year quarters:
       Q1 = Apr-Jun
       Q2 = Jul-Sep
       Q3 = Oct-Dec
       Q4 = Jan-Mar
    """
    if month_num in (4, 5, 6):
        return f"Q1 {year} (Apr–Jun)"
    elif month_num in (7, 8, 9):
        return f"Q2 {year} (Jul–Sep)"
    elif month_num in (10, 11, 12):
        return f"Q3 {year} (Oct–Dec)"
    else:
        return f"Q4 {year - 1}-{year} (Jan–Mar)"


def discover_quarterly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial quarters."""
    monthly = discover_monthly_files(data_dir)
    quarters = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        quarter = get_quarter_label(month_num, year)
        if quarter not in quarters:
            quarters[quarter] = []

        quarters[quarter].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for quarter in quarters:
        quarters[quarter] = sorted(
            quarters[quarter],
            key=lambda x: (x["year"], x["month"])
        )

    return quarters


def get_half_yearly_label(month_num: int, year: int) -> str:
    """Financial year half-years:
       H1 = Apr-Sep
       H2 = Oct-Mar
    """
    if month_num in (4, 5, 6, 7, 8, 9):
        return f"H1 {year} (Apr–Sep)"
    else:
        fy_year = year if month_num in (10, 11, 12) else year - 1
        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"


def discover_half_yearly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial half-years."""
    monthly = discover_monthly_files(data_dir)
    half_years = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        hy_label = get_half_yearly_label(month_num, year)
        if hy_label not in half_years:
            half_years[hy_label] = []

        half_years[hy_label].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for hy in half_years:
        half_years[hy] = sorted(
            half_years[hy],
            key=lambda x: (x["year"], x["month"])
        )

    return half_years


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-MONTH VRM AND DRM COMBINE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_selected_vrm_data(selected_entries: list) -> dict:
    """
    Loads VRM data for multiple selected month entries and concatenates 
    them together across 'active', 'dropped', and 'new_reg' sheets.
    """
    active_dfs, dropped_dfs, new_reg_dfs = [], [], []

    for entry in selected_entries:
        vrm_path = entry.get("vrm")
        if vrm_path and os.path.exists(vrm_path):
            month_data = load_data(vrm_path)
            if month_data:
                if "active" in month_data and not month_data["active"].empty:
                    df_act = month_data["active"].copy()
                    df_act["Reporting Period"] = entry["label"]
                    active_dfs.append(df_act)
                
                if "dropped" in month_data and not month_data["dropped"].empty:
                    df_drp = month_data["dropped"].copy()
                    df_drp["Reporting Period"] = entry["label"]
                    dropped_dfs.append(df_drp)
                
                if "new_reg" in month_data and not month_data["new_reg"].empty:
                    df_nr = month_data["new_reg"].copy()
                    df_nr["Reporting Period"] = entry["label"]
                    new_reg_dfs.append(df_nr)

    return {
        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),
        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),
        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),
    }


def load_selected_drm_data(selected_entries: list) -> dict:
    """
    Loads DRM data for multiple selected month entries and concatenates 
    them across all individual DRM sheets.
    """
    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []

    for entry in selected_entries:
        drm_path = entry.get("drm")
        if drm_path and os.path.exists(drm_path):
            m_drm = load_drm_data(drm_path)
            if m_drm:
                for key, container in [
                    ("sess", sess_dfs),
                    ("drm", drm_dfs),
                    ("ac", ac_dfs),
                    ("od", od_dfs),
                    ("ne", ne_dfs),
                ]:
                    if key in m_drm and not m_drm[key].empty:
                        df_sheet = m_drm[key].copy()
                        df_sheet["Reporting Period"] = entry["label"]
                        container.append(df_sheet)

    return {
        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),
        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),
        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),
        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),
        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),
    }


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

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

    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))
    dropped.columns = dropped.columns.str.strip()
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

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

DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

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

    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))
    ac.columns = ac.columns.str.strip()

    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

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

    # ─────────────────────────────────────────────────────────────────────────
    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)
    # ─────────────────────────────────────────────────────────────────────────

    reporting_type = st.radio(
        "📊 Reporting Type",
        ["Monthly", "Quarterly", "Half-Yearly"],
        horizontal=True,
        key="reporting_type",
    )

    selected_entries = []
    period_label = ""

    if reporting_type == "Monthly":
        month_labels = list(available_months.keys())

        selected_months = st.multiselect(
            "📅 Select Reporting Months",
            options=month_labels,
            default=[month_labels[-1]] if month_labels else [],
            key="selected_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            {
                "label": month,
                "vrm": available_months[month]["vrm"],
                "drm": available_months[month]["drm"],
            }
            for month in selected_months
        ]

        period_label = ", ".join(selected_months)

    elif reporting_type == "Quarterly":
        available_quarters = discover_quarterly_files(DATA_DIR)

        if not available_quarters:
            st.error("⚠️ No quarterly data available.")
            return

        quarter_labels = list(available_quarters.keys())

        selected_quarter = st.selectbox(
            "📅 Reporting Quarter",
            options=quarter_labels,
            index=len(quarter_labels) - 1,
            key="selected_quarter",
        )

        quarter_entries = available_quarters[selected_quarter]

        month_options = [
            item["label"]
            for item in quarter_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_quarter_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in quarter_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_quarter} — "
            + ", ".join(selected_months)
        )

    else:
        available_half_years = discover_half_yearly_files(DATA_DIR)

        if not available_half_years:
            st.error("⚠️ No half-yearly data available.")
            return

        hy_labels = list(available_half_years.keys())

        selected_hy = st.selectbox(
            "📅 Reporting Half-Year",
            options=hy_labels,
            index=len(hy_labels) - 1,
            key="selected_half_year",
        )

        hy_entries = available_half_years[selected_hy]

        month_options = [
            item["label"]
            for item in hy_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_hy_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in hy_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_hy} — "
            + ", ".join(selected_months)
        )

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {period_label}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD VRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading VRM data for {period_label}…"):
        data = load_selected_vrm_data(selected_entries)

    if not data or data["active"].empty:
        st.error(f"⚠️ Could not load VRM data for {period_label}.")
        return

    df_raw  = data["active"]
    dropped = data["dropped"]
    new_reg = data["new_reg"]

    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD DRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading DRM data for {period_label}…"):
        drm_data_raw = load_selected_drm_data(selected_entries)

    # ── Build sidebar — content switches based on active_tab flag ─────────────
    with st.sidebar:
        st.markdown("---")

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

        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())
            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

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

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    vol_df = df.drop_duplicates(subset="Volunteer id").copy() if not df.empty else df.copy()

    # ─────────────────────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────────────────────
    tab_vol, tab_ctr, tab_aca, tab_drm, tab_top, tab_bl = st.tabs([
        "🙋 Volunteers",
        "🏫 Centres",
        "📚 Academic Health",
        "📊 DRM Client Report",
        "📚 Topic & SubTopic Analytics",
        "🎯 Baseline Only",
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
        else:
            total_vols_offering = len(df)
            unique_vols         = vol_df["Volunteer id"].nunique()
            dropped_vols        = len(dropped)
            new_vols            = new_reg["User ID"].nunique() if "User ID" in new_reg.columns else 0
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
            if "Date Joined" in new_reg.columns and "User ID" in new_reg.columns:
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
            if "Gender" in new_reg.columns and "User ID" in new_reg.columns:
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
        else:
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
                st.info("🔍 Gender breakdown (En Boys / En Girls) isn't available in this export.")

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
        else:
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
    # TAB 4 — DRM CLIENT REPORT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:
        if not drm_data_raw or "sess" not in drm_data_raw or drm_data_raw["sess"].empty:
            st.error(f"⚠️ No DRM file found or DRM data is empty for **{period_label}**.")
        else:
            sess_raw = drm_data_raw["sess"]
            drm_raw  = drm_data_raw["drm"]
            od_raw   = drm_data_raw["od"]

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

            if sess.empty or drm.empty:
                st.markdown("<br>", unsafe_allow_html=True)
                st.info(
                    "🔍 **No data available for the selected filters.**\n\n"
                    "The combination of Donor, State, and Subject you selected returned no sessions "
                    "or no active centres. Try broadening your selection in the sidebar."
                )
            else:
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
                                    &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {period_label}
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

                total_sessions = len(sess)
                comp_sessions  = int((sess["Session_status"] == "Completed").sum())
                off_sessions   = int((sess["Session_status"] == "Offline").sum())
                canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
                total_clh_drm  = int(sess["Present/CLH"].sum())
                total_students = int(drm["Enrolled"].sum())
                avg_att_drm    = sess["Attendance%"].mean()
                total_centres  = drm["Center Name"].nunique()
                drm_total_vols = sess["Volunteer_id"].nunique() if "Volunteer_id" in sess.columns else 0
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

                for col_w, lbl, val, color, sub in [
                    (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
                    (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique active"),
                    (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned in period"),
                    (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
                    (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

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
                    (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{(off_sessions/total_sessions*100) if total_sessions else 0:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("---")

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

                st.markdown("#### ❌ Cancellations & CLH Split")

                col_ca1, col_ca2 = st.columns([1, 1.4])

                with col_ca1:
                    st.markdown("##### Cancellation reasons")
                    if "Cancel_reason" in sess.columns:
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
                    else:
                        st.info("Cancellation reasons are not available in this dataset.")

                with col_ca2:
                    st.markdown("##### State-wise CLH split")
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

                st.markdown("#### 📚 Subject & Grade Analytics")

                col_sg1, col_sg2 = st.columns(2)

                with col_sg1:
                    st.markdown("##### Sessions by subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
                        subj_grade_sess = (
                            sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                            .count().reset_index()
                            .rename(columns={"Session_id": "Sessions"})
                        )
                        pivot_sgs = (
                            subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                            .fillna(0)
                        )
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

                with col_sg2:
                    st.markdown("##### Session completion % — subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
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
                        if 'pivot_sgs' in locals():
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

                st.markdown("#### ⏱️ CLH by Subject & Grade")

                col_clh1, col_clh2 = st.columns([1.4, 1])

                with col_clh1:
                    st.markdown("##### CLH heatmap — subject × grade")
                    if "Grade" in sess.columns:
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

                st.markdown("#### 🙋 Volunteer Performance Scorecard")

                if {"Volunteer_id", "Teacher_name", "Session_id"}.issubset(sess.columns):
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

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                            padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                        eVidyaloka Programme Impact Report &nbsp;·&nbsp; {period_label}
                        &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                        {drm_total_vols} volunteers &nbsp;·&nbsp;
                        {total_students:,} students &nbsp;·&nbsp;
                        {total_clh_drm:,} child learning hours
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — TOPIC & SUBTOPIC ANALYTICS (INTEGRATED TOPIC DASHBOARD)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("### 📚 Topic & SubTopic Analytics Dashboard")

        uploaded_topic_file = st.file_uploader(
            "📁 Upload Topic/SubTopic Excel File",
            type=["xlsx", "xls"],
            key="topic_file_uploader",
            help="Upload the Excel file containing 'Topic-SubTopic', 'Cancelled Sessions', and 'Offline Sessions' sheets."
        )

        if uploaded_topic_file is None:
            st.info("👆 Please upload a Topic/SubTopic Excel file above to activate this dashboard.")
        else:
            @st.cache_data
            def load_topic_uploaded_data(file):
                try:
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names

                    def _get_sheet_name(target):
                        norm_target = " ".join(str(target).strip().lower().split())
                        for sheet in sheets:
                            if " ".join(str(sheet).strip().lower().split()) == norm_target:
                                return sheet
                        return target

                    t_sheet = _get_sheet_name("Topic-SubTopic")
                    c_sheet = _get_sheet_name("Cancelled Sessions")
                    o_sheet = _get_sheet_name("Offline Sessions")

                    t_df = pd.read_excel(file, sheet_name=t_sheet) if t_sheet in sheets else pd.DataFrame()
                    c_df = pd.read_excel(file, sheet_name=c_sheet) if c_sheet in sheets else pd.DataFrame()
                    o_df = pd.read_excel(file, sheet_name=o_sheet) if o_sheet in sheets else pd.DataFrame()

                    return t_df, c_df, o_df
                except Exception as e:
                    st.error(f"Error reading uploaded file: {e}")
                    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            topic_df, cancelled_df, offline_df = load_topic_uploaded_data(uploaded_topic_file)

            if topic_df.empty:
                st.error("The uploaded file does not contain a valid 'Topic-SubTopic' sheet or has no data.")
            else:
                # ── IN-TAB FILTERS ──────────────────────────────────────────
                with st.expander("📊 Topic Filters", expanded=True):
                    fc1, fc2, fc3 = st.columns(3)

                    state_col   = "state" if "state" in topic_df.columns else "State"
                    subject_col = "subject" if "subject" in topic_df.columns else "Subject"
                    grade_col   = "grade" if "grade" in topic_df.columns else "Grade"

                    with fc1:
                        st_opts = sorted(topic_df[state_col].dropna().unique()) if state_col in topic_df.columns else []
                        t_state_filter = st.multiselect("State", st_opts, key="tab5_state")

                    state_df = topic_df[topic_df[state_col].isin(t_state_filter)] if t_state_filter and state_col in topic_df.columns else topic_df.copy()

                    with fc2:
                        sub_opts = sorted(state_df[subject_col].dropna().unique()) if subject_col in state_df.columns else []
                        t_subject_filter = st.multiselect("Subject", sub_opts, key="tab5_subject")

                    subject_df = state_df[state_df[subject_col].isin(t_subject_filter)] if t_subject_filter and subject_col in state_df.columns else state_df.copy()

                    with fc3:
                        grd_opts = sorted(subject_df[grade_col].dropna().unique()) if grade_col in subject_df.columns else []
                        t_grade_filter = st.multiselect("Grade", grd_opts, key="tab5_grade")

                    filtered_topic_df = subject_df[subject_df[grade_col].isin(t_grade_filter)] if t_grade_filter and grade_col in subject_df.columns else subject_df.copy()

                # ── KPI METRICS ──────────────────────────────────────────────
                total_sessions   = len(filtered_topic_df)
                unique_topics    = filtered_topic_df["topic_name"].nunique() if "topic_name" in filtered_topic_df.columns else 0
                unique_subtopics = filtered_topic_df["sub_topic_name"].nunique() if "sub_topic_name" in filtered_topic_df.columns else 0
                unique_centers   = filtered_topic_df["center_id"].nunique() if "center_id" in filtered_topic_df.columns else (
                    filtered_topic_df["center_name"].nunique() if "center_name" in filtered_topic_df.columns else 0
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Sessions", f"{total_sessions:,}")
                with col2:
                    st.metric("Unique Topics", f"{unique_topics:,}")
                with col3:
                    st.metric("Unique Subtopics", f"{unique_subtopics:,}")
                with col4:
                    st.metric("Active Centers", f"{unique_centers:,}")

                st.markdown("---")

                # ── DASHBOARD SUB-TABS ────────────────────────────────────────
                top_tab1, top_tab2, top_tab3 = st.tabs([
                    "📋 Topic Analysis",
                    "❌ Cancelled Sessions",
                    "⏸️ Offline Sessions"
                ])

                # 1. TOPIC ANALYSIS
                with top_tab1:
                    if "topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Topics")
                        topic_summary = (
                            filtered_topic_df
                            .groupby("topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig = px.bar(
                            topic_summary, x="Sessions", y="topic_name",
                            orientation="h", title="Top 20 Topics"
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    if "sub_topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Sub Topics")
                        subtopic_summary = (
                            filtered_topic_df
                            .groupby("sub_topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig2 = px.bar(
                            subtopic_summary, x="Sessions", y="sub_topic_name",
                            orientation="h", title="Top 20 Sub Topics"
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    if subject_col in filtered_topic_df.columns:
                        st.subheader("Subject Distribution")
                        subject_chart = (
                            filtered_topic_df
                            .groupby(subject_col)
                            .size()
                            .reset_index(name="Count")
                        )
                        fig3 = px.pie(
                            subject_chart, names=subject_col, values="Count"
                        )
                        st.plotly_chart(fig3, use_container_width=True)

                # 2. CANCELLED SESSIONS
                with top_tab2:
                    st.subheader("Cancelled Session Analysis")
                    if not cancelled_df.empty:
                        st.metric("Total Cancelled Sessions", f"{len(cancelled_df):,}")

                        cancel_reason_col = (
                            "cancel_reason" if "cancel_reason" in cancelled_df.columns else (
                                "Cancel_reason" if "Cancel_reason" in cancelled_df.columns else None
                            )
                        )
                        if cancel_reason_col:
                            cancel_chart = (
                                cancelled_df
                                .groupby(cancel_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                                .head(15)
                            )
                            fig4 = px.bar(
                                cancel_chart, x="Count", y=cancel_reason_col,
                                orientation="h", title="Cancellation Reasons"
                            )
                            st.plotly_chart(fig4, use_container_width=True)

                        canc_state_col = "State" if "State" in cancelled_df.columns else (
                            "state" if "state" in cancelled_df.columns else None
                        )
                        if canc_state_col:
                            state_cancel = (
                                cancelled_df
                                .groupby(canc_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig5 = px.bar(
                                state_cancel, x=canc_state_col, y="Count",
                                title="Cancelled Sessions by State"
                            )
                            st.plotly_chart(fig5, use_container_width=True)
                    else:
                        st.info("No cancelled sessions data found in the uploaded file.")

                # 3. OFFLINE SESSIONS
                with top_tab3:
                    st.subheader("Offline Session Analysis")
                    if not offline_df.empty:
                        st.metric("Total Offline Sessions", f"{len(offline_df):,}")

                        offline_reason_col = (
                            "Offline reason" if "Offline reason" in offline_df.columns else (
                                "offline_reason" if "offline_reason" in offline_df.columns else None
                            )
                        )
                        if offline_reason_col:
                            offline_reason = (
                                offline_df
                                .groupby(offline_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig6 = px.bar(
                                offline_reason, x="Count", y=offline_reason_col,
                                orientation="h", title="Offline Reasons"
                            )
                            st.plotly_chart(fig6, use_container_width=True)

                        off_state_col = "State" if "State" in offline_df.columns else (
                            "state" if "state" in offline_df.columns else None
                        )
                        if off_state_col:
                            state_offline = (
                                offline_df
                                .groupby(off_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig7 = px.bar(
                                state_offline, x=off_state_col, y="Count",
                                title="Offline Sessions by State"
                            )
                            st.plotly_chart(fig7, use_container_width=True)
                    else:
                        st.info("No offline sessions data found in the uploaded file.")

                # ── EXPORT DATA ──────────────────────────────────────────────
                st.markdown("---")
                st.download_button(
                    "📥 Download Filtered Topic Data",
                    filtered_topic_df.to_csv(index=False),
                    file_name=f"filtered_{uploaded_topic_file.name.replace('.xlsx', '.csv')}",
                    mime="text/csv",
                    key="download_topic_csv"
                )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 6 — BASELINE ONLY METRICS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_bl:
        st.markdown("### 🎯 Baseline Assessment Diagnostics (BL Only)")
        st.caption(
            "Upload a standalone Baseline assessment Excel file to inspect student readiness, "
            "proficiency tiers, and subject/grade gaps without requiring Endline data."
        )

        bl_file = st.file_uploader(
            "📁 Upload Baseline Assessment File (.xlsx)",
            type=["xlsx", "xls"],
            key="bl_only_uploader",
        )

        if bl_file is None:
            st.info("👆 Upload your Baseline Excel file above to view Baseline-only metrics and visuals.")
        else:
            @st.cache_data(show_spinner=False)
            def _load_bl_only(file):
                xl = pd.ExcelFile(file)
                target_sheet = xl.sheet_names[0]
                for s in xl.sheet_names:
                    if any(k in s.lower() for k in ("baseline", "bl", "score", "data")):
                        target_sheet = s
                        break
                df_b = pd.read_excel(xl, sheet_name=target_sheet)
                df_b.columns = df_b.columns.str.strip()
                return df_b

            bl_raw = _load_bl_only(bl_file)

            if bl_raw.empty:
                st.error("⚠️ The uploaded Baseline sheet is empty.")
            else:
                def _pick_col(df_check, candidates):
                    lower_map = {c.lower().strip(): c for c in df_check.columns}
                    for cand in candidates:
                        if cand.lower() in lower_map:
                            return lower_map[cand.lower()]
                    return None

                col_score   = _pick_col(bl_raw, [
                    "Baseline%", "Baseline %", "BL%", "BL %",
                    "Baseline Score", "Score%", "Score %", "Percentage", "Marks %"
                ])
                col_student = _pick_col(bl_raw, ["Student ID", "Student_ID", "Student Name", "Child ID", "Enrollment ID"])
                col_center  = _pick_col(bl_raw, ["Center Name", "Center", "Centre Name", "Centre"])
                col_state   = _pick_col(bl_raw, ["State", "Center State"])
                col_subject = _pick_col(bl_raw, ["Subject", "Subject Name"])
                col_grade   = _pick_col(bl_raw, ["Grade", "Class"])
                col_gender  = _pick_col(bl_raw, ["Gender", "Sex"])

                if col_score is None:
                    st.error(
                        f"⚠️ Could not detect a percentage score column. Found columns: `{list(bl_raw.columns)}`. "
                        "Ensure a column like `Baseline%`, `BL %`, or `Percentage` exists."
                    )
                else:
                    bl_df = bl_raw.copy()
                    bl_df["_score"] = pd.to_numeric(bl_df[col_score], errors="coerce")
                    bl_df = bl_df.dropna(subset=["_score"])

                    # ── In-Tab Filters ───────────────────────────────────────
                    with st.expander("🔍 Baseline Filters", expanded=True):
                        f1, f2, f3 = st.columns(3)
                        with f1:
                            if col_state:
                                st_opts = sorted(bl_df[col_state].dropna().astype(str).unique())
                                sel_bl_state = st.multiselect("State", st_opts, default=st_opts, key="bl_only_state")
                                if sel_bl_state:
                                    bl_df = bl_df[bl_df[col_state].astype(str).isin(sel_bl_state)]
                        with f2:
                            if col_subject:
                                sub_opts = sorted(bl_df[col_subject].dropna().astype(str).unique())
                                sel_bl_sub = st.multiselect("Subject", sub_opts, default=sub_opts, key="bl_only_sub")
                                if sel_bl_sub:
                                    bl_df = bl_df[bl_df[col_subject].astype(str).isin(sel_bl_sub)]
                        with f3:
                            if col_grade:
                                grd_opts = sorted(bl_df[col_grade].dropna().astype(str).unique())
                                sel_bl_grd = st.multiselect("Grade", grd_opts, default=grd_opts, key="bl_only_grd")
                                if sel_bl_grd:
                                    bl_df = bl_df[bl_df[col_grade].astype(str).isin(sel_bl_grd)]

                    if bl_df.empty:
                        st.info("No records match the selected Baseline filters.")
                    else:
                        def _tier(v):
                            if v < 40:
                                return "Needs Support (<40%)"
                            elif v < 70:
                                return "Developing (40–69%)"
                            return "Proficient (≥70%)"

                        bl_df["Proficiency Tier"] = bl_df["_score"].apply(_tier)

                        # ── KPI Strip ────────────────────────────────────────
                        n_assessed  = bl_df[col_student].nunique() if col_student else len(bl_df)
                        avg_bl      = bl_df["_score"].mean()
                        med_bl      = bl_df["_score"].median()
                        prof_pct    = (bl_df["_score"] >= 70).mean() * 100
                        support_pct = (bl_df["_score"] < 40).mean() * 100
                        n_centres   = bl_df[col_center].nunique() if col_center else 0

                        k1, k2, k3, k4, k5 = st.columns(5)
                        k1.markdown(_kpi("Students Assessed", f"{n_assessed:,}",   P["teal"],   f"{n_centres} centres" if col_center else "total rows"), unsafe_allow_html=True)
                        k2.markdown(_kpi("Avg BL Score",      f"{avg_bl:.1f}%",     P["violet"], f"median {med_bl:.1f}%"), unsafe_allow_html=True)
                        k3.markdown(_kpi("Proficient (≥70%)", f"{prof_pct:.1f}%",   P["green"],  "ready / strong"), unsafe_allow_html=True)
                        k4.markdown(_kpi("Needs Support",     f"{support_pct:.1f}%",P["red"],    "below 40%"), unsafe_allow_html=True)
                        k5.markdown(_kpi("Total Assessments", f"{len(bl_df):,}",    P["amber"],  "subject-student tests"), unsafe_allow_html=True)

                        st.markdown("---")

                        # ── Row 1: Proficiency Donut + Score Distribution ────
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("##### Proficiency Tier Breakdown")
                            tier_df = bl_df["Proficiency Tier"].value_counts().reset_index()
                            tier_df.columns = ["Tier", "Count"]
                            fig_tier = px.pie(
                                tier_df, names="Tier", values="Count", hole=0.48,
                                color="Tier",
                                color_discrete_map={
                                    "Needs Support (<40%)": P["red"],
                                    "Developing (40–69%)":  P["amber"],
                                    "Proficient (≥70%)":    P["green"],
                                },
                            )
                            fig_tier.update_traces(
                                textposition="outside",
                                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                            )
                            fig_tier.update_layout(
                                height=320, showlegend=False,
                                plot_bgcolor=BG, paper_bgcolor=BG,
                                margin=dict(l=10, r=10, t=30, b=10),
                            )
                            st.plotly_chart(fig_tier, use_container_width=True)

                        with c2:
                            st.markdown("##### Baseline Score Distribution")
                            fig_bl_hist = px.histogram(
                                bl_df, x="_score", nbins=20,
                                color_discrete_sequence=[P["teal"]], opacity=0.85,
                            )
                            fig_bl_hist.add_vline(
                                x=avg_bl, line_dash="dash", line_color=P["red"],
                                annotation_text=f"Mean {avg_bl:.1f}%"
                            )
                            _layout(fig_bl_hist, height=320)
                            fig_bl_hist.update_layout(
                                xaxis_title="Baseline Score %",
                                yaxis_title="Assessments",
                                bargap=0.05,
                            )
                            st.plotly_chart(fig_bl_hist, use_container_width=True)

                        st.markdown("---")

                        # ── Row 2: Subject & Grade Performance ───────────────
                        c3, c4 = st.columns(2)
                        with c3:
                            if col_subject:
                                st.markdown("##### Avg Baseline % by Subject")
                                sub_bl = (
                                    bl_df.groupby(col_subject)["_score"].mean()
                                    .reset_index().rename(columns={"_score": "Avg BL%"})
                                    .sort_values("Avg BL%", ascending=True)
                                )
                                fig_sub_bl = px.bar(
                                    sub_bl, x="Avg BL%", y=col_subject, orientation="h",
                                    text=sub_bl["Avg BL%"].round(1).astype(str) + "%",
                                    color="Avg BL%",
                                    color_continuous_scale=[P["coral"], P["amber"], P["green"]],
                                    range_color=[0, 100],
                                )
                                fig_sub_bl.update_traces(textposition="outside", marker_line_width=0)
                                fig_sub_bl.update_coloraxes(showscale=False)
                                _layout(fig_sub_bl, height=320, margin=dict(l=0, r=60, t=20, b=0))
                                fig_sub_bl.update_layout(
                                    xaxis=dict(range=[0, 112], title="Avg Baseline %"),
                                    yaxis_title="",
                                )
                                st.plotly_chart(fig_sub_bl, use_container_width=True)

                        with c4:
                            if col_subject and col_grade:
                                st.markdown("##### Subject × Grade Avg Baseline %")
                                sg_piv = bl_df.pivot_table(
                                    index=col_subject, columns=col_grade,
                                    values="_score", aggfunc="mean"
                                ).fillna(0)
                                fig_sg_bl = px.imshow(
                                    sg_piv, text_auto=".0f", aspect="auto",
                                    color_continuous_scale=["#fadbd8", P["amber"], P["teal"]],
                                    zmin=0, zmax=100,
                                )
                                fig_sg_bl.update_layout(
                                    height=320, plot_bgcolor=BG, paper_bgcolor=BG,
                                    margin=dict(l=0, r=0, t=20, b=0),
                                    coloraxis_showscale=False,
                                )
                                st.plotly_chart(fig_sg_bl, use_container_width=True)

                        # ── Row 3: State & Gender Diagnostics ────────────────
                        if col_state or col_gender:
                            st.markdown("---")
                            c5, c6 = st.columns([1.4, 1])
                            with c5:
                                if col_state:
                                    st.markdown("##### State-wise Avg Baseline %")
                                    st_bl = (
                                        bl_df.groupby(col_state)["_score"].mean()
                                        .reset_index().rename(columns={"_score": "Avg BL%"})
                                        .sort_values("Avg BL%", ascending=True)
                                    )
                                    fig_st_bl = px.bar(
                                        st_bl, x="Avg BL%", y=col_state, orientation="h",
                                        text=st_bl["Avg BL%"].round(1).astype(str) + "%",
                                        color_discrete_sequence=[P["teal"]],
                                    )
                                    fig_st_bl.update_traces(textposition="outside", marker_line_width=0)
                                    _layout(fig_st_bl, height=max(260, len(st_bl) * 36), margin=dict(l=0, r=60, t=20, b=0))
                                    fig_st_bl.update_layout(
                                        xaxis=dict(range=[0, 112], title="Avg Baseline %"),
                                        yaxis_title="",
                                    )
                                    st.plotly_chart(fig_st_bl, use_container_width=True)
                            with c6:
                                if col_gender and col_subject:
                                    st.markdown("##### Gender Comparison by Subject")
                                    gen_bl = bl_df.groupby([col_subject, col_gender])["_score"].mean().reset_index()
                                    fig_gen_bl = px.bar(
                                        gen_bl, x=col_subject, y="_score", color=col_gender,
                                        barmode="group", text_auto=".1f",
                                        color_discrete_map={
                                            "Male": P["teal"], "Female": P["coral"],
                                            "Boys": P["teal"], "Girls": P["coral"],
                                        },
                                    )
                                    _layout(fig_gen_bl, height=300, legend_bottom=True)
                                    fig_gen_bl.update_layout(
                                        yaxis=dict(range=[0, 105], title="Avg BL %"),
                                        xaxis_title="",
                                    )
                                    st.plotly_chart(fig_gen_bl, use_container_width=True)

                        # ── Priority Centres Table ───────────────────────────
                        if col_center:
                            st.markdown("---")
                            st.markdown("##### 🏫 Centre Baseline Leaderboard (Lowest to Highest — Priority Support)")
                            ctr_grp_cols = [col_center] + ([col_state] if col_state else [])
                            ctr_bl = (
                                bl_df.groupby(ctr_grp_cols).agg(
                                    Assessments=("_score", "count"),
                                    Avg_BL_Pct=("_score", "mean"),
                                    Needs_Support_Pct=("_score", lambda s: (s < 40).mean() * 100),
                                    Proficient_Pct=("_score", lambda s: (s >= 70).mean() * 100),
                                ).reset_index()
                                .sort_values("Avg_BL_Pct", ascending=True)
                            )
                            st.dataframe(
                                ctr_bl,
                                use_container_width=True,
                                hide_index=True,
                                height=380,
                                column_config={
                                    "Avg_BL_Pct": st.column_config.ProgressColumn(
                                        "Avg BL %", min_value=0, max_value=100, format="%.1f%%"
                                    ),
                                    "Needs_Support_Pct": st.column_config.NumberColumn(
                                        "Needs Support %", format="%.1f%%"
                                    ),
                                    "Proficient_Pct": st.column_config.NumberColumn(
                                        "Proficient %", format="%.1f%%"
                                    ),
                                },
                            )
error code File "/mount/src/impact-analytics-dashboard/app.py", line 10, in <module>

from ops_dashboard import show_dashboard

File "/mount/src/impact-analytics-dashboard/ops_dashboard.py", line 1942, in <module>

with tab_topic:

^^^^^^^^^

Two distinct issues are triggering this traceback:

Indentation error at line 1942 (with tab_topic:): In ops_dashboard.py, the code starting around with tab_topic: (and any tabs below it) was pasted at the top module level (0 spaces of indentation) instead of inside def render_ops_dashboard(): (4 spaces of indentation). Because Python executes module-level statements immediately on import, it hits with tab_topic: before render_ops_dashboard() runs and throws NameError: name 'tab_topic' is not defined.   
ops_dashboard (15).py
+ 1

Function name mismatch in app.py (from ops_dashboard import show_dashboard): Your ops_dashboard.py defines def render_ops_dashboard():, whereas app.py is importing show_dashboard. Adding a show_dashboard = render_ops_dashboard alias at the bottom of ops_dashboard.py fixes that ImportError automatically without needing to touch app.py.   
ops_dashboard (15).py

How to Fix Both Errors
1. Fix the Indentation in ops_dashboard.py
Open ops_dashboard.py around line 1940 and make sure with tab_topic: (and every line inside it, as well as with tab_bl:) is indented 4 spaces so it sits inside def render_ops_dashboard(): aligned with with tab_drm: and with tab_top:.   
ops_dashboard (15).py

Python
    # Aligned at 4 spaces inside render_ops_dashboard():
    with tab_drm:
        ...

    with tab_top:   # <-- Ensure this matches your st.tabs(...) variable name!
        ...

    with tab_bl:
        ...
(Note: Check your st.tabs(...) unpack line too — in the file you shared earlier it was named tab_top, while your traceback shows tab_topic. Make sure the variable name in tab_vol, tab_ctr, tab_aca, tab_drm, tab_topic, tab_bl = st.tabs([...]) matches the with tab_...: blocks.)   
ops_dashboard (15).py

2. Add the show_dashboard Alias at the Bottom of ops_dashboard.py
At the very end of ops_dashboard.py (at 0 spaces indentation, outside render_ops_dashboard()), add:

Python
# Alias so `from ops_dashboard import show_dashboard` in app.py works seamlessly
show_dashboard = render_ops_dashboard
Verify the Fix
Save ops_dashboard.py and push/restart your Streamlit app.

Confirm the app loads without NameError or ImportError on startup, and verify that all tabs (including Topic Analytics and Baseline Only) appear in the horizontal tab bar.

If you want me to patch the exact lines cleanly for you, paste lines 1920 to the end of your current ops_dashboard.py.

can you share completed updated and corrected file

Your updated Python file is ready

PY icon
ops_dashboard
 PY 
Summary of Fixes & Updates in ops_dashboard.py
Syntax & Indentation Verified (py_compile passed): All 6 tabs (Volunteers, Centres, Academic Health, DRM Client Report, Topic & SubTopic Analytics, and Baseline Only) are properly nested inside render_ops_dashboard().

Resolved ImportError in app.py: Added show_dashboard = render_ops_dashboard at the bottom of the file so from ops_dashboard import show_dashboard works directly.

Non-blocking Tab Rendering: Replaced internal st.stop() calls inside individual tabs with if/else guards so empty filter selections in one tab never prevent subsequent tabs from rendering.

Added 🎯 Baseline Only Tab: Includes automatic column detection, State/Subject/Grade filters, 5 KPI cards, Proficiency Tier breakdown, Score Distribution histogram, Subject × Grade heatmap, Gender comparison, and a Priority Support Centre leaderboard.





Your eVidyaloka Trust chats aren't used to improve our models. Gemini is AI and can make mistakes. Your privacy and GeminiOpens in a new window


"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: auto-detected monthly workbooks named VRM_<Month>_<Year>.xlsx /
DRM_<Month>_<Year>.xlsx (e.g. VRM_April_2026.xlsx), selected via Monthly,
Quarterly, or Half-Yearly reporting modes. See discover_monthly_files().

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
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

# Matches VRM_April_2026.xlsx, DRM_May_2026.xlsx, VRM__May__2026.xlsx, etc.
MONTH_FILE_RE = re.compile(r"^(VRM|DRM)_+([A-Za-z]+)_+(\d{4})\.xlsx$", re.IGNORECASE)

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
    them by reporting month.
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


# ─────────────────────────────────────────────────────────────────────────────
# QUARTERLY & HALF-YEARLY REPORTING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_quarter_label(month_num: int, year: int) -> str:
    """Financial year quarters:
       Q1 = Apr-Jun
       Q2 = Jul-Sep
       Q3 = Oct-Dec
       Q4 = Jan-Mar
    """
    if month_num in (4, 5, 6):
        return f"Q1 {year} (Apr–Jun)"
    elif month_num in (7, 8, 9):
        return f"Q2 {year} (Jul–Sep)"
    elif month_num in (10, 11, 12):
        return f"Q3 {year} (Oct–Dec)"
    else:
        return f"Q4 {year - 1}-{year} (Jan–Mar)"


def discover_quarterly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial quarters."""
    monthly = discover_monthly_files(data_dir)
    quarters = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        quarter = get_quarter_label(month_num, year)
        if quarter not in quarters:
            quarters[quarter] = []

        quarters[quarter].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for quarter in quarters:
        quarters[quarter] = sorted(
            quarters[quarter],
            key=lambda x: (x["year"], x["month"])
        )

    return quarters


def get_half_yearly_label(month_num: int, year: int) -> str:
    """Financial year half-years:
       H1 = Apr-Sep
       H2 = Oct-Mar
    """
    if month_num in (4, 5, 6, 7, 8, 9):
        return f"H1 {year} (Apr–Sep)"
    else:
        fy_year = year if month_num in (10, 11, 12) else year - 1
        return f"H2 {fy_year}-{fy_year + 1} (Oct–Mar)"


def discover_half_yearly_files(data_dir: str) -> dict:
    """Group available monthly VRM/DRM files into financial half-years."""
    monthly = discover_monthly_files(data_dir)
    half_years = {}

    for label, paths in monthly.items():
        month_name, year = label.rsplit(" ", 1)
        month_num = _month_sort_key(month_name)
        year = int(year)

        if month_num == 0:
            continue

        hy_label = get_half_yearly_label(month_num, year)
        if hy_label not in half_years:
            half_years[hy_label] = []

        half_years[hy_label].append({
            "label": label,
            "month": month_num,
            "year": year,
            "vrm": paths["vrm"],
            "drm": paths["drm"],
        })

    for hy in half_years:
        half_years[hy] = sorted(
            half_years[hy],
            key=lambda x: (x["year"], x["month"])
        )

    return half_years


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-MONTH VRM AND DRM COMBINE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_selected_vrm_data(selected_entries: list) -> dict:
    """
    Loads VRM data for multiple selected month entries and concatenates 
    them together across 'active', 'dropped', and 'new_reg' sheets.
    """
    active_dfs, dropped_dfs, new_reg_dfs = [], [], []

    for entry in selected_entries:
        vrm_path = entry.get("vrm")
        if vrm_path and os.path.exists(vrm_path):
            month_data = load_data(vrm_path)
            if month_data:
                if "active" in month_data and not month_data["active"].empty:
                    df_act = month_data["active"].copy()
                    df_act["Reporting Period"] = entry["label"]
                    active_dfs.append(df_act)
                
                if "dropped" in month_data and not month_data["dropped"].empty:
                    df_drp = month_data["dropped"].copy()
                    df_drp["Reporting Period"] = entry["label"]
                    dropped_dfs.append(df_drp)
                
                if "new_reg" in month_data and not month_data["new_reg"].empty:
                    df_nr = month_data["new_reg"].copy()
                    df_nr["Reporting Period"] = entry["label"]
                    new_reg_dfs.append(df_nr)

    return {
        "active": pd.concat(active_dfs, ignore_index=True) if active_dfs else pd.DataFrame(),
        "dropped": pd.concat(dropped_dfs, ignore_index=True) if dropped_dfs else pd.DataFrame(),
        "new_reg": pd.concat(new_reg_dfs, ignore_index=True) if new_reg_dfs else pd.DataFrame(),
    }


def load_selected_drm_data(selected_entries: list) -> dict:
    """
    Loads DRM data for multiple selected month entries and concatenates 
    them across all individual DRM sheets.
    """
    sess_dfs, drm_dfs, ac_dfs, od_dfs, ne_dfs = [], [], [], [], []

    for entry in selected_entries:
        drm_path = entry.get("drm")
        if drm_path and os.path.exists(drm_path):
            m_drm = load_drm_data(drm_path)
            if m_drm:
                for key, container in [
                    ("sess", sess_dfs),
                    ("drm", drm_dfs),
                    ("ac", ac_dfs),
                    ("od", od_dfs),
                    ("ne", ne_dfs),
                ]:
                    if key in m_drm and not m_drm[key].empty:
                        df_sheet = m_drm[key].copy()
                        df_sheet["Reporting Period"] = entry["label"]
                        container.append(df_sheet)

    return {
        "sess": pd.concat(sess_dfs, ignore_index=True) if sess_dfs else pd.DataFrame(),
        "drm": pd.concat(drm_dfs, ignore_index=True) if drm_dfs else pd.DataFrame(),
        "ac": pd.concat(ac_dfs, ignore_index=True) if ac_dfs else pd.DataFrame(),
        "od": pd.concat(od_dfs, ignore_index=True) if od_dfs else pd.DataFrame(),
        "ne": pd.concat(ne_dfs, ignore_index=True) if ne_dfs else pd.DataFrame(),
    }


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    def _norm(s: str) -> str:
        return " ".join(str(s).strip().lower().split())

    target_norm = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_norm:
            return name
    return target


def _apply_aliases(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

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

    dropped = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_DROPPED))
    dropped.columns = dropped.columns.str.strip()
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    new_reg = pd.read_excel(xl, sheet_name=_find_sheet(xl, SHEET_NEW_REG))
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}

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

DRM_SHEET_ALIASES = {
    "Donor Name":     ["Donor"],
    "Live_CLH":       ["Live CLH"],
    "Attendance %":   ["Live Attendance %"],
}


@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}

    xl = pd.ExcelFile(path)

    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

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

    ac = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Active centers"))
    ac.columns = ac.columns.str.strip()

    od = pd.read_excel(xl, sheet_name=_find_sheet(xl, "Offering Details"))
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

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

    # ─────────────────────────────────────────────────────────────────────────
    # REPORTING PERIOD (MONTHLY VS QUARTERLY VS HALF-YEARLY MULTI-SELECT)
    # ─────────────────────────────────────────────────────────────────────────

    reporting_type = st.radio(
        "📊 Reporting Type",
        ["Monthly", "Quarterly", "Half-Yearly"],
        horizontal=True,
        key="reporting_type",
    )

    selected_entries = []
    period_label = ""

    if reporting_type == "Monthly":
        month_labels = list(available_months.keys())

        selected_months = st.multiselect(
            "📅 Select Reporting Months",
            options=month_labels,
            default=[month_labels[-1]] if month_labels else [],
            key="selected_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            {
                "label": month,
                "vrm": available_months[month]["vrm"],
                "drm": available_months[month]["drm"],
            }
            for month in selected_months
        ]

        period_label = ", ".join(selected_months)

    elif reporting_type == "Quarterly":
        available_quarters = discover_quarterly_files(DATA_DIR)

        if not available_quarters:
            st.error("⚠️ No quarterly data available.")
            return

        quarter_labels = list(available_quarters.keys())

        selected_quarter = st.selectbox(
            "📅 Reporting Quarter",
            options=quarter_labels,
            index=len(quarter_labels) - 1,
            key="selected_quarter",
        )

        quarter_entries = available_quarters[selected_quarter]

        month_options = [
            item["label"]
            for item in quarter_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_quarter_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in quarter_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_quarter} — "
            + ", ".join(selected_months)
        )

    else:
        available_half_years = discover_half_yearly_files(DATA_DIR)

        if not available_half_years:
            st.error("⚠️ No half-yearly data available.")
            return

        hy_labels = list(available_half_years.keys())

        selected_hy = st.selectbox(
            "📅 Reporting Half-Year",
            options=hy_labels,
            index=len(hy_labels) - 1,
            key="selected_half_year",
        )

        hy_entries = available_half_years[selected_hy]

        month_options = [
            item["label"]
            for item in hy_entries
        ]

        selected_months = st.multiselect(
            "📅 Select Months",
            options=month_options,
            default=month_options,
            key="selected_hy_months",
            placeholder="Select one or more months",
        )

        if not selected_months:
            st.warning("⚠️ Please select at least one month.")
            return

        selected_entries = [
            item
            for item in hy_entries
            if item["label"] in selected_months
        ]

        period_label = (
            f"{selected_hy} — "
            + ", ".join(selected_months)
        )

    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        f"  —  {period_label}</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD VRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading VRM data for {period_label}…"):
        data = load_selected_vrm_data(selected_entries)

    if not data or data["active"].empty:
        st.error(f"⚠️ Could not load VRM data for {period_label}.")
        return

    df_raw  = data["active"]
    dropped = data["dropped"]
    new_reg = data["new_reg"]

    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "vrm"

    # ═════════════════════════════════════════════════════════════════════════
    # LOAD DRM DATA
    # ═════════════════════════════════════════════════════════════════════════

    with st.spinner(f"Loading DRM data for {period_label}…"):
        drm_data_raw = load_selected_drm_data(selected_entries)

    # ── Build sidebar — content switches based on active_tab flag ─────────────
    with st.sidebar:
        st.markdown("---")

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

        if st.session_state.active_tab == "vrm":
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0094c9;margin-bottom:8px;'>🎯 VRM Filters</div>",
                unsafe_allow_html=True,
            )

            _vrm_all_donors   = sorted(df_raw["Donor"].dropna().unique())
            _vrm_all_states   = sorted(df_raw["State"].dropna().unique())
            _vrm_all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

            sel_donors = st.multiselect(
                "Donor",
                options=_vrm_all_donors,
                default=_vrm_all_donors,
                placeholder="Type to search donors…",
                key="ops_donor",
            )

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

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
                sel_drm_donors   = sorted(drm_data_raw["sess"]["Donor"].dropna().unique())
                sel_drm_states   = sorted(drm_data_raw["sess"]["State"].dropna().unique())
                sel_drm_subjects = sorted(drm_data_raw["sess"]["Subject_clean"].dropna().unique())
            else:
                sel_drm_donors = sel_drm_states = sel_drm_subjects = []

        else:
            st.markdown(
                "<div style='font-size:0.75em;font-weight:700;text-transform:uppercase;"
                "letter-spacing:0.07em;color:#0f8a6e;margin-bottom:8px;'>📊 DRM Filters</div>",
                unsafe_allow_html=True,
            )

            if drm_data_raw and "sess" in drm_data_raw and not drm_data_raw["sess"].empty:
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

            sel_donors   = sorted(df_raw["Donor"].dropna().unique())
            sel_states   = sorted(df_raw["State"].dropna().unique())
            sel_subjects = sorted(df_raw["Subject_clean"].dropna().unique())

    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    vol_df = df.drop_duplicates(subset="Volunteer id").copy() if not df.empty else df.copy()

    # ─────────────────────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────────────────────
    tab_vol, tab_ctr, tab_aca, tab_drm, tab_top, tab_bl = st.tabs([
        "🙋 Volunteers",
        "🏫 Centres",
        "📚 Academic Health",
        "📊 DRM Client Report",
        "📚 Topic & SubTopic Analytics",
        "🎯 Baseline Only",
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
        else:
            total_vols_offering = len(df)
            unique_vols         = vol_df["Volunteer id"].nunique()
            dropped_vols        = len(dropped)
            new_vols            = new_reg["User ID"].nunique() if "User ID" in new_reg.columns else 0
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
            if "Date Joined" in new_reg.columns and "User ID" in new_reg.columns:
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
            if "Gender" in new_reg.columns and "User ID" in new_reg.columns:
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
        else:
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
                st.info("🔍 Gender breakdown (En Boys / En Girls) isn't available in this export.")

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
        else:
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
    # TAB 4 — DRM CLIENT REPORT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_drm:
        if not drm_data_raw or "sess" not in drm_data_raw or drm_data_raw["sess"].empty:
            st.error(f"⚠️ No DRM file found or DRM data is empty for **{period_label}**.")
        else:
            sess_raw = drm_data_raw["sess"]
            drm_raw  = drm_data_raw["drm"]
            od_raw   = drm_data_raw["od"]

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

            if sess.empty or drm.empty:
                st.markdown("<br>", unsafe_allow_html=True)
                st.info(
                    "🔍 **No data available for the selected filters.**\n\n"
                    "The combination of Donor, State, and Subject you selected returned no sessions "
                    "or no active centres. Try broadening your selection in the sidebar."
                )
            else:
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
                                    &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; {period_label}
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

                total_sessions = len(sess)
                comp_sessions  = int((sess["Session_status"] == "Completed").sum())
                off_sessions   = int((sess["Session_status"] == "Offline").sum())
                canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())
                total_clh_drm  = int(sess["Present/CLH"].sum())
                total_students = int(drm["Enrolled"].sum())
                avg_att_drm    = sess["Attendance%"].mean()
                total_centres  = drm["Center Name"].nunique()
                drm_total_vols = sess["Volunteer_id"].nunique() if "Volunteer_id" in sess.columns else 0
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

                for col_w, lbl, val, color, sub in [
                    (r1c1, "Active Centres",    f"{total_centres}",      D["teal"],   f"across {num_states} states"),
                    (r1c2, "Active Volunteers", f"{drm_total_vols}",     D["green"],  "unique active"),
                    (r1c3, "Total Sessions",    f"{total_sessions:,}",   D["blue"],   "planned in period"),
                    (r1c4, "Completed",         f"{comp_sessions:,}",    D["teal"],   f"{comp_rate:.1f}% rate"),
                    (r1c5, "Cancelled",         f"{canc_sessions}",      D["coral"],  f"{canc_rate:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

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
                    (r2c5, "Offline Sessions",f"{off_sessions}",      D["amber"],  f"{(off_sessions/total_sessions*100) if total_sessions else 0:.1f}% of planned"),
                ]:
                    col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("---")

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

                st.markdown("#### ❌ Cancellations & CLH Split")

                col_ca1, col_ca2 = st.columns([1, 1.4])

                with col_ca1:
                    st.markdown("##### Cancellation reasons")
                    if "Cancel_reason" in sess.columns:
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
                    else:
                        st.info("Cancellation reasons are not available in this dataset.")

                with col_ca2:
                    st.markdown("##### State-wise CLH split")
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

                st.markdown("#### 📚 Subject & Grade Analytics")

                col_sg1, col_sg2 = st.columns(2)

                with col_sg1:
                    st.markdown("##### Sessions by subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
                        subj_grade_sess = (
                            sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                            .count().reset_index()
                            .rename(columns={"Session_id": "Sessions"})
                        )
                        pivot_sgs = (
                            subj_grade_sess.pivot(index="Subject_clean", columns="Grade", values="Sessions")
                            .fillna(0)
                        )
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

                with col_sg2:
                    st.markdown("##### Session completion % — subject & grade")
                    if "Grade" in sess.columns and "Session_id" in sess.columns:
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
                        if 'pivot_sgs' in locals():
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

                st.markdown("#### ⏱️ CLH by Subject & Grade")

                col_clh1, col_clh2 = st.columns([1.4, 1])

                with col_clh1:
                    st.markdown("##### CLH heatmap — subject × grade")
                    if "Grade" in sess.columns:
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

                st.markdown("#### 🙋 Volunteer Performance Scorecard")

                if {"Volunteer_id", "Teacher_name", "Session_id"}.issubset(sess.columns):
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

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                            padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                        eVidyaloka Programme Impact Report &nbsp;·&nbsp; {period_label}
                        &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                        {drm_total_vols} volunteers &nbsp;·&nbsp;
                        {total_students:,} students &nbsp;·&nbsp;
                        {total_clh_drm:,} child learning hours
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 5 — TOPIC & SUBTOPIC ANALYTICS (INTEGRATED TOPIC DASHBOARD)
    # ═════════════════════════════════════════════════════════════════════════
    with tab_top:
        st.markdown("### 📚 Topic & SubTopic Analytics Dashboard")

        uploaded_topic_file = st.file_uploader(
            "📁 Upload Topic/SubTopic Excel File",
            type=["xlsx", "xls"],
            key="topic_file_uploader",
            help="Upload the Excel file containing 'Topic-SubTopic', 'Cancelled Sessions', and 'Offline Sessions' sheets."
        )

        if uploaded_topic_file is None:
            st.info("👆 Please upload a Topic/SubTopic Excel file above to activate this dashboard.")
        else:
            @st.cache_data
            def load_topic_uploaded_data(file):
                try:
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names

                    def _get_sheet_name(target):
                        norm_target = " ".join(str(target).strip().lower().split())
                        for sheet in sheets:
                            if " ".join(str(sheet).strip().lower().split()) == norm_target:
                                return sheet
                        return target

                    t_sheet = _get_sheet_name("Topic-SubTopic")
                    c_sheet = _get_sheet_name("Cancelled Sessions")
                    o_sheet = _get_sheet_name("Offline Sessions")

                    t_df = pd.read_excel(file, sheet_name=t_sheet) if t_sheet in sheets else pd.DataFrame()
                    c_df = pd.read_excel(file, sheet_name=c_sheet) if c_sheet in sheets else pd.DataFrame()
                    o_df = pd.read_excel(file, sheet_name=o_sheet) if o_sheet in sheets else pd.DataFrame()

                    return t_df, c_df, o_df
                except Exception as e:
                    st.error(f"Error reading uploaded file: {e}")
                    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            topic_df, cancelled_df, offline_df = load_topic_uploaded_data(uploaded_topic_file)

            if topic_df.empty:
                st.error("The uploaded file does not contain a valid 'Topic-SubTopic' sheet or has no data.")
            else:
                # ── IN-TAB FILTERS ──────────────────────────────────────────
                with st.expander("📊 Topic Filters", expanded=True):
                    fc1, fc2, fc3 = st.columns(3)

                    state_col   = "state" if "state" in topic_df.columns else "State"
                    subject_col = "subject" if "subject" in topic_df.columns else "Subject"
                    grade_col   = "grade" if "grade" in topic_df.columns else "Grade"

                    with fc1:
                        st_opts = sorted(topic_df[state_col].dropna().unique()) if state_col in topic_df.columns else []
                        t_state_filter = st.multiselect("State", st_opts, key="tab5_state")

                    state_df = topic_df[topic_df[state_col].isin(t_state_filter)] if t_state_filter and state_col in topic_df.columns else topic_df.copy()

                    with fc2:
                        sub_opts = sorted(state_df[subject_col].dropna().unique()) if subject_col in state_df.columns else []
                        t_subject_filter = st.multiselect("Subject", sub_opts, key="tab5_subject")

                    subject_df = state_df[state_df[subject_col].isin(t_subject_filter)] if t_subject_filter and subject_col in state_df.columns else state_df.copy()

                    with fc3:
                        grd_opts = sorted(subject_df[grade_col].dropna().unique()) if grade_col in subject_df.columns else []
                        t_grade_filter = st.multiselect("Grade", grd_opts, key="tab5_grade")

                    filtered_topic_df = subject_df[subject_df[grade_col].isin(t_grade_filter)] if t_grade_filter and grade_col in subject_df.columns else subject_df.copy()

                # ── KPI METRICS ──────────────────────────────────────────────
                total_sessions   = len(filtered_topic_df)
                unique_topics    = filtered_topic_df["topic_name"].nunique() if "topic_name" in filtered_topic_df.columns else 0
                unique_subtopics = filtered_topic_df["sub_topic_name"].nunique() if "sub_topic_name" in filtered_topic_df.columns else 0
                unique_centers   = filtered_topic_df["center_id"].nunique() if "center_id" in filtered_topic_df.columns else (
                    filtered_topic_df["center_name"].nunique() if "center_name" in filtered_topic_df.columns else 0
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Sessions", f"{total_sessions:,}")
                with col2:
                    st.metric("Unique Topics", f"{unique_topics:,}")
                with col3:
                    st.metric("Unique Subtopics", f"{unique_subtopics:,}")
                with col4:
                    st.metric("Active Centers", f"{unique_centers:,}")

                st.markdown("---")

                # ── DASHBOARD SUB-TABS ────────────────────────────────────────
                top_tab1, top_tab2, top_tab3 = st.tabs([
                    "📋 Topic Analysis",
                    "❌ Cancelled Sessions",
                    "⏸️ Offline Sessions"
                ])

                # 1. TOPIC ANALYSIS
                with top_tab1:
                    if "topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Topics")
                        topic_summary = (
                            filtered_topic_df
                            .groupby("topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig = px.bar(
                            topic_summary, x="Sessions", y="topic_name",
                            orientation="h", title="Top 20 Topics"
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    if "sub_topic_name" in filtered_topic_df.columns:
                        st.subheader("Top Sub Topics")
                        subtopic_summary = (
                            filtered_topic_df
                            .groupby("sub_topic_name")
                            .size()
                            .reset_index(name="Sessions")
                            .sort_values("Sessions", ascending=False)
                            .head(20)
                        )
                        fig2 = px.bar(
                            subtopic_summary, x="Sessions", y="sub_topic_name",
                            orientation="h", title="Top 20 Sub Topics"
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    if subject_col in filtered_topic_df.columns:
                        st.subheader("Subject Distribution")
                        subject_chart = (
                            filtered_topic_df
                            .groupby(subject_col)
                            .size()
                            .reset_index(name="Count")
                        )
                        fig3 = px.pie(
                            subject_chart, names=subject_col, values="Count"
                        )
                        st.plotly_chart(fig3, use_container_width=True)

                # 2. CANCELLED SESSIONS
                with top_tab2:
                    st.subheader("Cancelled Session Analysis")
                    if not cancelled_df.empty:
                        st.metric("Total Cancelled Sessions", f"{len(cancelled_df):,}")

                        cancel_reason_col = (
                            "cancel_reason" if "cancel_reason" in cancelled_df.columns else (
                                "Cancel_reason" if "Cancel_reason" in cancelled_df.columns else None
                            )
                        )
                        if cancel_reason_col:
                            cancel_chart = (
                                cancelled_df
                                .groupby(cancel_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                                .head(15)
                            )
                            fig4 = px.bar(
                                cancel_chart, x="Count", y=cancel_reason_col,
                                orientation="h", title="Cancellation Reasons"
                            )
                            st.plotly_chart(fig4, use_container_width=True)

                        canc_state_col = "State" if "State" in cancelled_df.columns else (
                            "state" if "state" in cancelled_df.columns else None
                        )
                        if canc_state_col:
                            state_cancel = (
                                cancelled_df
                                .groupby(canc_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig5 = px.bar(
                                state_cancel, x=canc_state_col, y="Count",
                                title="Cancelled Sessions by State"
                            )
                            st.plotly_chart(fig5, use_container_width=True)
                    else:
                        st.info("No cancelled sessions data found in the uploaded file.")

                # 3. OFFLINE SESSIONS
                with top_tab3:
                    st.subheader("Offline Session Analysis")
                    if not offline_df.empty:
                        st.metric("Total Offline Sessions", f"{len(offline_df):,}")

                        offline_reason_col = (
                            "Offline reason" if "Offline reason" in offline_df.columns else (
                                "offline_reason" if "offline_reason" in offline_df.columns else None
                            )
                        )
                        if offline_reason_col:
                            offline_reason = (
                                offline_df
                                .groupby(offline_reason_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig6 = px.bar(
                                offline_reason, x="Count", y=offline_reason_col,
                                orientation="h", title="Offline Reasons"
                            )
                            st.plotly_chart(fig6, use_container_width=True)

                        off_state_col = "State" if "State" in offline_df.columns else (
                            "state" if "state" in offline_df.columns else None
                        )
                        if off_state_col:
                            state_offline = (
                                offline_df
                                .groupby(off_state_col)
                                .size()
                                .reset_index(name="Count")
                                .sort_values("Count", ascending=False)
                            )
                            fig7 = px.bar(
                                state_offline, x=off_state_col, y="Count",
                                title="Offline Sessions by State"
                            )
                            st.plotly_chart(fig7, use_container_width=True)
                    else:
                        st.info("No offline sessions data found in the uploaded file.")

                # ── EXPORT DATA ──────────────────────────────────────────────
                st.markdown("---")
                st.download_button(
                    "📥 Download Filtered Topic Data",
                    filtered_topic_df.to_csv(index=False),
                    file_name=f"filtered_{uploaded_topic_file.name.replace('.xlsx', '.csv')}",
                    mime="text/csv",
                    key="download_topic_csv"
                )

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 6 — BASELINE ONLY METRICS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_bl:
        st.markdown("### 🎯 Baseline Assessment Diagnostics (BL Only)")
        st.caption(
            "Upload a standalone Baseline assessment Excel file to inspect student readiness, "
            "proficiency tiers, and subject/grade gaps without requiring Endline data."
        )

        bl_file = st.file_uploader(
            "📁 Upload Baseline Assessment File (.xlsx)",
            type=["xlsx", "xls"],
            key="bl_only_uploader",
        )

        if bl_file is None:
            st.info("👆 Upload your Baseline Excel file above to view Baseline-only metrics and visuals.")
        else:
            @st.cache_data(show_spinner=False)
            def _load_bl_only(file):
                xl = pd.ExcelFile(file)
                target_sheet = xl.sheet_names[0]
                for s in xl.sheet_names:
                    if any(k in s.lower() for k in ("baseline", "bl", "score", "data")):
                        target_sheet = s
                        break
                df_b = pd.read_excel(xl, sheet_name=target_sheet)
                df_b.columns = df_b.columns.str.strip()
                return df_b

            bl_raw = _load_bl_only(bl_file)

            if bl_raw.empty:
                st.error("⚠️ The uploaded Baseline sheet is empty.")
            else:
                def _pick_col(df_check, candidates):
                    lower_map = {c.lower().strip(): c for c in df_check.columns}
                    for cand in candidates:
                        if cand.lower() in lower_map:
                            return lower_map[cand.lower()]
                    return None

                col_score   = _pick_col(bl_raw, [
                    "Baseline%", "Baseline %", "BL%", "BL %",
                    "Baseline Score", "Score%", "Score %", "Percentage", "Marks %"
                ])
                col_student = _pick_col(bl_raw, ["Student ID", "Student_ID", "Student Name", "Child ID", "Enrollment ID"])
                col_center  = _pick_col(bl_raw, ["Center Name", "Center", "Centre Name", "Centre"])
                col_state   = _pick_col(bl_raw, ["State", "Center State"])
                col_subject = _pick_col(bl_raw, ["Subject", "Subject Name"])
                col_grade   = _pick_col(bl_raw, ["Grade", "Class"])
                col_gender  = _pick_col(bl_raw, ["Gender", "Sex"])

                if col_score is None:
                    st.error(
                        f"⚠️ Could not detect a percentage score column. Found columns: `{list(bl_raw.columns)}`. "
                        "Ensure a column like `Baseline%`, `BL %`, or `Percentage` exists."
                    )
                else:
                    bl_df = bl_raw.copy()
                    bl_df["_score"] = pd.to_numeric(bl_df[col_score], errors="coerce")
                    bl_df = bl_df.dropna(subset=["_score"])

                    # ── In-Tab Filters ───────────────────────────────────────
                    with st.expander("🔍 Baseline Filters", expanded=True):
                        f1, f2, f3 = st.columns(3)
                        with f1:
                            if col_state:
                                st_opts = sorted(bl_df[col_state].dropna().astype(str).unique())
                                sel_bl_state = st.multiselect("State", st_opts, default=st_opts, key="bl_only_state")
                                if sel_bl_state:
                                    bl_df = bl_df[bl_df[col_state].astype(str).isin(sel_bl_state)]
                        with f2:
                            if col_subject:
                                sub_opts = sorted(bl_df[col_subject].dropna().astype(str).unique())
                                sel_bl_sub = st.multiselect("Subject", sub_opts, default=sub_opts, key="bl_only_sub")
                                if sel_bl_sub:
                                    bl_df = bl_df[bl_df[col_subject].astype(str).isin(sel_bl_sub)]
                        with f3:
                            if col_grade:
                                grd_opts = sorted(bl_df[col_grade].dropna().astype(str).unique())
                                sel_bl_grd = st.multiselect("Grade", grd_opts, default=grd_opts, key="bl_only_grd")
                                if sel_bl_grd:
                                    bl_df = bl_df[bl_df[col_grade].astype(str).isin(sel_bl_grd)]

                    if bl_df.empty:
                        st.info("No records match the selected Baseline filters.")
                    else:
                        def _tier(v):
                            if v < 40:
                                return "Needs Support (<40%)"
                            elif v < 70:
                                return "Developing (40–69%)"
                            return "Proficient (≥70%)"

                        bl_df["Proficiency Tier"] = bl_df["_score"].apply(_tier)

                        # ── KPI Strip ────────────────────────────────────────
                        n_assessed  = bl_df[col_student].nunique() if col_student else len(bl_df)
                        avg_bl      = bl_df["_score"].mean()
                        med_bl      = bl_df["_score"].median()
                        prof_pct    = (bl_df["_score"] >= 70).mean() * 100
                        support_pct = (bl_df["_score"] < 40).mean() * 100
                        n_centres   = bl_df[col_center].nunique() if col_center else 0

                        k1, k2, k3, k4, k5 = st.columns(5)
                        k1.markdown(_kpi("Students Assessed", f"{n_assessed:,}",   P["teal"],   f"{n_centres} centres" if col_center else "total rows"), unsafe_allow_html=True)
                        k2.markdown(_kpi("Avg BL Score",      f"{avg_bl:.1f}%",     P["violet"], f"median {med_bl:.1f}%"), unsafe_allow_html=True)
                        k3.markdown(_kpi("Proficient (≥70%)", f"{prof_pct:.1f}%",   P["green"],  "ready / strong"), unsafe_allow_html=True)
                        k4.markdown(_kpi("Needs Support",     f"{support_pct:.1f}%",P["red"],    "below 40%"), unsafe_allow_html=True)
                        k5.markdown(_kpi("Total Assessments", f"{len(bl_df):,}",    P["amber"],  "subject-student tests"), unsafe_allow_html=True)

                        st.markdown("---")

                        # ── Row 1: Proficiency Donut + Score Distribution ────
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("##### Proficiency Tier Breakdown")
                            tier_df = bl_df["Proficiency Tier"].value_counts().reset_index()
                            tier_df.columns = ["Tier", "Count"]
                            fig_tier = px.pie(
                                tier_df, names="Tier", values="Count", hole=0.48,
                                color="Tier",
                                color_discrete_map={
                                    "Needs Support (<40%)": P["red"],
                                    "Developing (40–69%)":  P["amber"],
                                    "Proficient (≥70%)":    P["green"],
                                },
                            )
                            fig_tier.update_traces(
                                textposition="outside",
                                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                            )
                            fig_tier.update_layout(
                                height=320, showlegend=False,
                                plot_bgcolor=BG, paper_bgcolor=BG,
                                margin=dict(l=10, r=10, t=30, b=10),
                            )
                            st.plotly_chart(fig_tier, use_container_width=True)

                        with c2:
                            st.markdown("##### Baseline Score Distribution")
                            fig_bl_hist = px.histogram(
                                bl_df, x="_score", nbins=20,
                                color_discrete_sequence=[P["teal"]], opacity=0.85,
                            )
                            fig_bl_hist.add_vline(
                                x=avg_bl, line_dash="dash", line_color=P["red"],
                                annotation_text=f"Mean {avg_bl:.1f}%"
                            )
                            _layout(fig_bl_hist, height=320)
                            fig_bl_hist.update_layout(
                                xaxis_title="Baseline Score %",
                                yaxis_title="Assessments",
                                bargap=0.05,
                            )
                            st.plotly_chart(fig_bl_hist, use_container_width=True)

                        st.markdown("---")

                        # ── Row 2: Subject & Grade Performance ───────────────
                        c3, c4 = st.columns(2)
                        with c3:
                            if col_subject:
                                st.markdown("##### Avg Baseline % by Subject")
                                sub_bl = (
                                    bl_df.groupby(col_subject)["_score"].mean()
                                    .reset_index().rename(columns={"_score": "Avg BL%"})
                                    .sort_values("Avg BL%", ascending=True)
                                )
                                fig_sub_bl = px.bar(
                                    sub_bl, x="Avg BL%", y=col_subject, orientation="h",
                                    text=sub_bl["Avg BL%"].round(1).astype(str) + "%",
                                    color="Avg BL%",
                                    color_continuous_scale=[P["coral"], P["amber"], P["green"]],
                                    range_color=[0, 100],
                                )
                                fig_sub_bl.update_traces(textposition="outside", marker_line_width=0)
                                fig_sub_bl.update_coloraxes(showscale=False)
                                _layout(fig_sub_bl, height=320, margin=dict(l=0, r=60, t=20, b=0))
                                fig_sub_bl.update_layout(
                                    xaxis=dict(range=[0, 112], title="Avg Baseline %"),
                                    yaxis_title="",
                                )
                                st.plotly_chart(fig_sub_bl, use_container_width=True)

                        with c4:
                            if col_subject and col_grade:
                                st.markdown("##### Subject × Grade Avg Baseline %")
                                sg_piv = bl_df.pivot_table(
                                    index=col_subject, columns=col_grade,
                                    values="_score", aggfunc="mean"
                                ).fillna(0)
                                fig_sg_bl = px.imshow(
                                    sg_piv, text_auto=".0f", aspect="auto",
                                    color_continuous_scale=["#fadbd8", P["amber"], P["teal"]],
                                    zmin=0, zmax=100,
                                )
                                fig_sg_bl.update_layout(
                                    height=320, plot_bgcolor=BG, paper_bgcolor=BG,
                                    margin=dict(l=0, r=0, t=20, b=0),
                                    coloraxis_showscale=False,
                                )
                                st.plotly_chart(fig_sg_bl, use_container_width=True)

                        # ── Row 3: State & Gender Diagnostics ────────────────
                        if col_state or col_gender:
                            st.markdown("---")
                            c5, c6 = st.columns([1.4, 1])
                            with c5:
                                if col_state:
                                    st.markdown("##### State-wise Avg Baseline %")
                                    st_bl = (
                                        bl_df.groupby(col_state)["_score"].mean()
                                        .reset_index().rename(columns={"_score": "Avg BL%"})
                                        .sort_values("Avg BL%", ascending=True)
                                    )
                                    fig_st_bl = px.bar(
                                        st_bl, x="Avg BL%", y=col_state, orientation="h",
                                        text=st_bl["Avg BL%"].round(1).astype(str) + "%",
                                        color_discrete_sequence=[P["teal"]],
                                    )
                                    fig_st_bl.update_traces(textposition="outside", marker_line_width=0)
                                    _layout(fig_st_bl, height=max(260, len(st_bl) * 36), margin=dict(l=0, r=60, t=20, b=0))
                                    fig_st_bl.update_layout(
                                        xaxis=dict(range=[0, 112], title="Avg Baseline %"),
                                        yaxis_title="",
                                    )
                                    st.plotly_chart(fig_st_bl, use_container_width=True)
                            with c6:
                                if col_gender and col_subject:
                                    st.markdown("##### Gender Comparison by Subject")
                                    gen_bl = bl_df.groupby([col_subject, col_gender])["_score"].mean().reset_index()
                                    fig_gen_bl = px.bar(
                                        gen_bl, x=col_subject, y="_score", color=col_gender,
                                        barmode="group", text_auto=".1f",
                                        color_discrete_map={
                                            "Male": P["teal"], "Female": P["coral"],
                                            "Boys": P["teal"], "Girls": P["coral"],
                                        },
                                    )
                                    _layout(fig_gen_bl, height=300, legend_bottom=True)
                                    fig_gen_bl.update_layout(
                                        yaxis=dict(range=[0, 105], title="Avg BL %"),
                                        xaxis_title="",
                                    )
                                    st.plotly_chart(fig_gen_bl, use_container_width=True)

                        # ── Priority Centres Table ───────────────────────────
                        if col_center:
                            st.markdown("---")
                            st.markdown("##### 🏫 Centre Baseline Leaderboard (Lowest to Highest — Priority Support)")
                            ctr_grp_cols = [col_center] + ([col_state] if col_state else [])
                            ctr_bl = (
                                bl_df.groupby(ctr_grp_cols).agg(
                                    Assessments=("_score", "count"),
                                    Avg_BL_Pct=("_score", "mean"),
                                    Needs_Support_Pct=("_score", lambda s: (s < 40).mean() * 100),
                                    Proficient_Pct=("_score", lambda s: (s >= 70).mean() * 100),
                                ).reset_index()
                                .sort_values("Avg_BL_Pct", ascending=True)
                            )
                            st.dataframe(
                                ctr_bl,
                                use_container_width=True,
                                hide_index=True,
                                height=380,
                                column_config={
                                    "Avg_BL_Pct": st.column_config.ProgressColumn(
                                        "Avg BL %", min_value=0, max_value=100, format="%.1f%%"
                                    ),
                                    "Needs_Support_Pct": st.column_config.NumberColumn(
                                        "Needs Support %", format="%.1f%%"
                                    ),
                                    "Proficient_Pct": st.column_config.NumberColumn(
                                        "Proficient %", format="%.1f%%"
                                    ),
                                },
                            )


# Alias so `from ops_dashboard import show_dashboard` in app.py works seamlessly
show_dashboard = render_ops_dashboard
ops_dashboard.py
Displaying ops_dashboard.py.
