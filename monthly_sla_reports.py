# ============================================================
# CELL 1 — IMPORTS
# ============================================================

import os
import smtplib
import numpy as np
import pandas as pd
import requests
import urllib3

from collections import defaultdict, deque
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================
# CELL 2 — CONFIGURATION
# ============================================================

# --- API ---
base_url = ""
headers = {
    "Authorization": "",
    "Content-Type": "application/json",
}

# --- Paths (Jupyter-safe) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(BASE_DIR, "region_reports")
FINAL_FILE = os.path.join(BASE_DIR, "final_table.xlsx")

# --- Email ---
REPORT_MONTH    = pd.Timestamp.today().strftime("%B %Y")
SMTP_SERVER     = ""
SMTP_PORT       = 
SENDER_EMAIL    = ""
SENDER_PASSWORD = ""
MY_EMAIL        = ""

DRY_RUN = False   # Set False when ready to actually send emails

# --- Staffing target ---
TARGET_SLA = 0.95

# --- People ---
SERVICE_HOD_NAME  = ""
SERVICE_HOD_EMAIL = ""
HR_NAME           = ""
HR_EMAIL          = ""

ASM_EMAIL_MAP = {
    
}

VALID_SUB_REGIONS = {
    
}

# Locked ASM -> sub-region ownership
ASM_REGION_MAP = {
    
}

# Job-type -> expected engineer role (contamination logic - DO NOT CHANGE)
JOB_EXPECTED_ROLE = {
    "PREVENTIVE MAINTENANCE VISIT": "Assurance",
    "INSPECTION VISIT":             "Service",
    "COMPLAINT":                    "Service",
    "SERVICE AND VISIT":            "Service",
    "BREAKDOWN VISIT":              "Service",
    "PVT ASSIST VISIT":             "Service",
    "COURTESY VISIT":               "Service",
    "ASTM VALIDATION":              "Service",
    "UPGRADATION":                  "Service",
    "INSTALLATION":                 "Service",
}

# Fixed business-defined SLA time limits per job type (in minutes)
# These are the agreed thresholds - what changes each month is job performance against them
JOB_SLA_LIMITS = {
    "INSTALLATION":                 240,
    "BREAKDOWN VISIT":              180,
    "PREVENTIVE MAINTENANCE VISIT": 300,
    "INSPECTION VISIT":             240,
    "SERVICE AND VISIT":            240,
    "COURTESY VISIT":               240,
    "UPGRADATION":                  240,
    "ASTM VALIDATION":              240,
    "COMPLAINT":                    180,
}

SERVICE_DESIGNATIONS   = ["Service Engineer", "Sr. Service Engineer"]
ASSURANCE_DESIGNATIONS = ["Assurance Service Engineer"]

# Territory -> sub-region mapping
TERRITORY_SUBREGION_OVERRIDE = {
    
}

# Month-over-month improvement direction
IMPROVE_RULES = {
    "additional_required_staffs":   "down",
    "total_required_staff_95":      "down",
    "contamination_pct":            "down",
    "jobs_late_pct":                "down",
    "jobs_on_time_comfortable_pct": "up",
    "jobs_on_time_near_sla_pct":    "up",
    "current_staff":                "up",
}

PCT_COLS = [
    "jobs_on_time_comfortable_pct",
    "jobs_on_time_near_sla_pct",
    "jobs_late_pct",
    "contamination_pct",
]

FINAL_COLUMNS = [
    "sub_region",
    "engineer_role",
    "current_staff",
    "total_required_staff_95",
    "additional_required_staffs",
    "jobs_on_time_comfortable_pct",
    "jobs_on_time_near_sla_pct",
    "jobs_late_pct",
    "contamination_pct",
    "contamination_status",
    "overall_status",
    "report_date",
]

# ============================================================
# CELL 3 — HELPER FUNCTIONS
# ============================================================

def paginated_get(endpoint, fields):
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=0.1,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))

    url = base_url + endpoint
    all_data, limit_start = [], 0
    while True:
        params = {
            "fields": fields,
            "limit_start": limit_start,
            "limit_page_length": 10000000000000,
        }
        try:
            response = session.get(
                url, headers=headers, params=params,
                timeout=120, verify=False
            )
            response.raise_for_status()
            data = response.json().get("data", [])
            if not data:
                break
            all_data.extend(data)
            limit_start += 100000000000000
        except requests.exceptions.RequestException as e:
            print(f"API Error ({endpoint}): {e}")
            break
    return all_data


def build_children_map(territory_df):
    children = defaultdict(list)
    for _, row in territory_df.iterrows():
        if pd.notna(row["parent"]):
            children[row["parent"]].append(row["territory"])
    return children


def get_descendants(root, children_map):
    descendants, queue = set(), deque([root])
    while queue:
        node = queue.popleft()
        for child in children_map.get(node, []):
            if child not in descendants:
                descendants.add(child)
                queue.append(child)
    return descendants


def map_engineer_role(designation):
    if designation in SERVICE_DESIGNATIONS:
        return "Service"
    if designation in ASSURANCE_DESIGNATIONS:
        return "Assurance"
    return None


def region_to_filename(region):
    return region.replace(" ", "_") + ".xlsx"


def contamination_improvement(row):
    curr = row.get("contamination_pct_curr")
    prev = row.get("contamination_pct_prev")
    
    if pd.isna(curr) or pd.isna(prev):
        return "Baseline"
    
    # Contamination should go DOWN (lower is better)
    # diff = prev - curr means positive = improvement
    diff = prev - curr
    
    if diff > 0.5:  # More than 0.5 percentage point improvement
        return f"Improved ({diff:+.1f}pp)"
    elif diff < -0.5:  # More than 0.5pp deterioration
        return f"Worsened ({diff:.1f}pp)"
    else:
        return "Stable"


def overall_improvement(row):
    diffs = []
    for col, direction in IMPROVE_RULES.items():
        curr = row.get(f"{col}_curr")
        prev = row.get(f"{col}_prev")
        if pd.isna(curr) or pd.isna(prev) or prev == 0:
            continue
        pct = ((curr - prev) / prev) * 100
        if direction == "down":
            pct = -pct
        diffs.append(pct)
    if not diffs:
        return "No Change"
    net = round(sum(diffs) / len(diffs), 1)
    if net > 0:
        return f"Improved (+{net}%)"
    elif net < 0:
        return f"Worsened ({net}%)"
    return "No Change"


# ============================================================
# CELL 4 — FETCH RAW DATA
# ============================================================

print("Fetching Service Reports...")
df_service = pd.DataFrame(paginated_get(
    "",
    '["job_type","item_name","completion_status","engineer_name","engineer_designation","territory","start_time","end_time"]',
))
print(f"  -> {len(df_service)} records")

print("Fetching Employees...")
df_employee_raw = pd.DataFrame(paginated_get(
    "",
    '["employee_name","status","designation","custom_territory"]',
))
print(f"  -> {len(df_employee_raw)} records")

print("Fetching Territory Hierarchy...")
territory_raw = pd.DataFrame(paginated_get(
    "",
    '["territory_name","parent_territory"]',
))
print(f"  -> {len(territory_raw)} records")


# ============================================================
# CELL 5 — BUILD TERRITORY HIERARCHY
# ============================================================

territory_df = territory_raw.rename(columns={
    "territory_name":   "territory",
    "parent_territory": "parent",
})

territory_df["parent"] = territory_df["parent"].replace(TERRITORY_SUBREGION_OVERRIDE)

territory_df = territory_df[
    ~territory_df["parent"].str.contains("Export", na=False) &
    (territory_df["territory"] != "All Territories")
].copy()

children_map = build_children_map(territory_df)

# Build territory -> sub_region lookup via BFS from each valid sub-region
territory_to_subregion = {}
for sub in VALID_SUB_REGIONS:
    territory_to_subregion[sub] = sub
    for t in get_descendants(sub, children_map):
        territory_to_subregion[t] = sub


# ============================================================
# CELL 6 — PREPARE SERVICE JOBS + DTC (DO NOT CHANGE)
# ============================================================

df_jobs = df_service.copy()
df_jobs["start_time"] = pd.to_datetime(df_jobs["start_time"], errors="coerce")
df_jobs["end_time"]   = pd.to_datetime(df_jobs["end_time"],   errors="coerce")

# DTC = Door to Completion in minutes
df_jobs["DTC_minutes"] = (
    (df_jobs["end_time"] - df_jobs["start_time"]).dt.total_seconds() / 60
)

# Remove invalid or absurd durations
df_jobs = df_jobs[
    (df_jobs["DTC_minutes"] > 0) &
    (df_jobs["DTC_minutes"] <= 24 * 60)
].copy()

# Map territory -> sub_region
df_jobs["territory_norm"] = df_jobs["territory"].astype(str).str.strip().str.title()

territory_map_clean = (
    territory_df[["territory", "parent"]]
    .rename(columns={"territory": "territory_norm_raw", "parent": "sub_region"})
    .assign(territory_norm=lambda x: x["territory_norm_raw"].str.strip().str.title())
    .drop_duplicates(subset=["territory_norm"])
    .dropna(subset=["sub_region"])
)
territory_map_clean = territory_map_clean[
    territory_map_clean["sub_region"].str.strip() != ""
]

df_jobs = df_jobs.merge(
    territory_map_clean[["territory_norm", "sub_region"]],
    on="territory_norm",
    how="left",
)
df_jobs = df_jobs.dropna(subset=["sub_region"]).reset_index(drop=True)
df_jobs = df_jobs[~df_jobs["sub_region"].isin(["India", "All Territories", "Export"])]

# Map engineer role
df_jobs["engineer_role"] = df_jobs["engineer_designation"].map(map_engineer_role)
df_jobs = df_jobs.dropna(subset=["engineer_role"]).reset_index(drop=True)

print(f"Service jobs after cleaning: {len(df_jobs)}")

# ============================================================
# CELL 7 — SLA BUCKETS (fixed business-defined limits per job type)
# ============================================================
# JOB_SLA_LIMITS are fixed business rules (e.g. BREAKDOWN VISIT = 180 min).
# What changes each month is how many jobs fell inside or outside these limits.
#   <= 50% of limit  -> On Time (Comfortable)
#   50% to 100%      -> On Time (Near SLA)
#   > 100%           -> Late (SLA Failed)

df_jobs["sla_time_limit_minutes"] = df_jobs["job_type"].map(JOB_SLA_LIMITS)

# Only keep jobs that have a defined SLA limit
df_sla = df_jobs.dropna(subset=["sla_time_limit_minutes"]).copy()

df_sla["sla_bucket"] = np.where(
    df_sla["DTC_minutes"] <= 0.5 * df_sla["sla_time_limit_minutes"],
    "On Time (Comfortable)",
    np.where(
        df_sla["DTC_minutes"] <= df_sla["sla_time_limit_minutes"],
        "On Time (Near SLA)",
        "Late (SLA Failed)",
    ),
)

# Quick sanity check - inspect a sample
print(df_sla["sla_bucket"].value_counts(normalize=True).mul(100).round(1))
print("\nSample rows:")
print(df_sla[["job_type", "DTC_minutes", "sla_time_limit_minutes", "sla_bucket"]].sample(10))


# ============================================================
# CELL 8 — CONTAMINATION LOGIC (DO NOT CHANGE)
# ============================================================

df_sla["expected_role"]   = df_sla["job_type"].map(JOB_EXPECTED_ROLE)
df_sla = df_sla.dropna(subset=["expected_role"])
df_sla["is_contaminated"] = df_sla["engineer_role"] != df_sla["expected_role"]

contamination = (
    df_sla
    .groupby(["sub_region", "expected_role"])
    .agg(
        total_jobs=("is_contaminated", "size"),
        contaminated_jobs=("is_contaminated", "sum"),
    )
    .reset_index()
)
contamination["contamination_pct"] = (
    contamination["contaminated_jobs"] / contamination["total_jobs"] * 100
).round(1)
contamination = contamination.rename(columns={"expected_role": "engineer_role"})

# ============================================================
# CELL 9 — SLA SUMMARY TABLE
# ============================================================

sla_summary = (
    df_sla
    .groupby(["sub_region", "engineer_role"])["sla_bucket"]
    .value_counts(normalize=True)
    .unstack(fill_value=0)
    .reset_index()
)

for col in sla_summary.columns:
    if col not in ["sub_region", "engineer_role"]:
        sla_summary[col] = (sla_summary[col] * 100).round(1)

sla_summary = sla_summary.rename(columns={
    "On Time (Comfortable)": "jobs_on_time_comfortable_pct",
    "On Time (Near SLA)":    "jobs_on_time_near_sla_pct",
    "Late (SLA Failed)":     "jobs_late_pct",
})

sla_summary["sla_pct"] = (1 - sla_summary["jobs_late_pct"] / 100).round(3)

# ============================================================
# CELL 10 — HEADCOUNT (BFS across territory hierarchy)
# ============================================================

df_employee = df_employee_raw[
    (df_employee_raw["status"] == "Active") &
    (df_employee_raw["designation"].isin(SERVICE_DESIGNATIONS + ASSURANCE_DESIGNATIONS))
].copy()

df_employee["engineer_role"] = df_employee["designation"].apply(map_engineer_role)
df_employee = df_employee.rename(columns={"custom_territory": "territory"})

staff_rows = []

for sub in VALID_SUB_REGIONS:
    descendants = get_descendants(sub, children_map)

    emp_all = df_employee[
        df_employee["territory"].isin(descendants.union({sub}))
    ]

    counts = emp_all.groupby("engineer_role").size().to_dict()

    for role in ["Service", "Assurance"]:
        staff_rows.append({
            "sub_region": sub,
            "engineer_role": role,
            "current_staff": counts.get(role, 0),
        })

staffing = pd.DataFrame(staff_rows)

# ============================================================
# CELL 11 — FINAL TABLE
# ============================================================

final_table = sla_summary.merge(
    staffing, on=["sub_region", "engineer_role"], how="left"
)

final_table["current_staff"] = final_table["current_staff"].fillna(0).astype(int)

# Safe SLA
final_table["sla_pct"] = final_table["sla_pct"].round(3)

# Required staff
final_table["total_required_staff_95"] = np.where(
    (final_table["current_staff"] > 0) & (final_table["sla_pct"] > 0),
    np.ceil((final_table["current_staff"] * TARGET_SLA) / final_table["sla_pct"]),
    0,
).astype(int)

# Minimum staffing rule
final_table.loc[
    (final_table["current_staff"] == 0) &
    (final_table["jobs_late_pct"].notna()),
    "total_required_staff_95"
] = 1

# Additional staff
final_table["additional_required_staffs"] = (
    final_table["total_required_staff_95"] - final_table["current_staff"]
).clip(lower=0).astype(int)

# Contamination
final_table = final_table.merge(
    contamination[["sub_region", "engineer_role", "contamination_pct"]],
    on=["sub_region", "engineer_role"],
    how="left",
)

final_table["contamination_pct"] = final_table["contamination_pct"].fillna(0)

# Round %
final_table[PCT_COLS] = final_table[PCT_COLS].round(2)

print(final_table.head(20))

# ============================================================
# CELL 12 — MONTH-OVER-MONTH STATUS
# ============================================================

if os.path.exists(FINAL_FILE):
    history_df = pd.read_excel(FINAL_FILE)
    history_df["report_date"] = pd.to_datetime(history_df["report_date"], errors="coerce")
    today     = pd.Timestamp.today().normalize()
    prev_data = history_df[history_df["report_date"] < today]

    if not prev_data.empty:
        latest_prev = (
            prev_data
            .sort_values("report_date")
            .groupby(["sub_region", "engineer_role"])
            .last()
            .reset_index()
        )
        compare_df = final_table.merge(
            latest_prev,
            on=["sub_region", "engineer_role"],
            how="left",
            suffixes=("_curr", "_prev"),
        )
        final_table["contamination_status"] = compare_df.apply(contamination_improvement, axis=1)
        final_table["overall_status"] = compare_df.apply(overall_improvement, axis=1)
    else:
        final_table["contamination_status"] = "Baseline"
        final_table["overall_status"] = "Baseline (First Month)"
else:
    final_table["contamination_status"] = "Baseline"
    final_table["overall_status"] = "Baseline (First Month)"

final_table["report_date"] = pd.Timestamp.today().date()

final_sla_table = (
    final_table[FINAL_COLUMNS]
    .sort_values(["sub_region", "engineer_role"])
    .reset_index(drop=True)
)

print(final_sla_table)

# ============================================================
# CELL 13 — APPEND TO MASTER FILE (monthly accumulation)
# ============================================================

if os.path.exists(FINAL_FILE):
    old_df      = pd.read_excel(FINAL_FILE)
    combined_df = pd.concat([old_df, final_sla_table], ignore_index=True)
else:
    combined_df = final_sla_table.copy()

combined_df["report_date"] = pd.to_datetime(combined_df["report_date"], errors="coerce").dt.date
combined_df = combined_df.drop_duplicates(
    subset=["sub_region", "engineer_role", "report_date"],
    keep="last",
)
combined_df = combined_df.sort_values(["report_date", "sub_region", "engineer_role"])
combined_df.to_excel(FINAL_FILE, index=False)

print(f"Master file updated: {FINAL_FILE}")
print(f"Total rows in file : {len(combined_df)}")

# ============================================================
# CELL 14 — GENERATE PER-REGION FILES (append history)
# ============================================================

os.makedirs(REPORT_DIR, exist_ok=True)
region_files = {}

for sr in final_sla_table["sub_region"].dropna().unique():
    region_df = final_sla_table[final_sla_table["sub_region"] == sr].copy()
    file_path = os.path.join(REPORT_DIR, region_to_filename(sr))

    if os.path.exists(file_path):
        old_region      = pd.read_excel(file_path)
        region_combined = pd.concat([old_region, region_df], ignore_index=True)
        region_combined["report_date"] = pd.to_datetime(
            region_combined["report_date"], errors="coerce"
        ).dt.date
        region_combined = region_combined.drop_duplicates(
            subset=["sub_region", "engineer_role", "report_date"],
            keep="last",
        )
    else:
        region_combined = region_df

    region_combined = region_combined.sort_values(["report_date", "engineer_role"])
    region_combined.to_excel(file_path, index=False)
    region_files[sr] = file_path
    print(f"Generated: {region_to_filename(sr)}")

# ============================================================
# CELL 15 — EMAIL DISPATCH
# ============================================================

asm_files_map = defaultdict(list)
for asm, regions in ASM_REGION_MAP.items():
    for region in regions:
        if region in region_files:
            asm_files_map[asm].append(region_files[region])
