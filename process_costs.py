"""
Azure FinOps — Monthly Cost Dashboard Generator
================================================
Usage:  python process_costs.py
Output: output/dashboard_YYYY-MM.html  (open in any browser, share by email/Teams/SharePoint)

Folder layout expected:
  finops-dashboard/
  ├── process_costs.py          ← this file
  ├── exports/                  ← drop Azure XLSX exports here (any filename, auto-detected)
  │   ├── monthly_april.xlsx
  │   ├── daily_april.xlsx
  │   └── resources_april.xlsx
  └── output/                   ← generated HTML files appear here (created automatically)

Each month needs up to 3 Azure Cost Management exports (all optional if missing):
  1. Monthly granularity  — Cost analysis → Group by: ResourceGroupName → Granularity: Monthly
  2. Daily granularity    — Cost analysis → Group by: ResourceGroupName → Granularity: Daily
  3. Resource-level       — Cost analysis → View: Resources (no grouping) → full month

How to export from Azure Portal:
  1. Go to portal.azure.com → Cost Management + Billing → Cost analysis
  2. Set scope to your subscription
  3. Set date range to the full month (e.g. Apr 01–Apr 30)
  4. For export 1: Group by = ResourceGroupName, Granularity = Monthly → Download → Excel
  5. For export 2: Group by = ResourceGroupName, Granularity = Daily  → Download → Excel
  6. For export 3: Switch to "Resources" view → Download → Excel
  7. Rename files clearly (e.g. monthly_2026_04.xlsx) and drop in exports/
"""

# ============================================================
# CONFIGURATION — edit this section to match your environment
# ============================================================

EXPORTS_DIR = "exports"        # folder containing Azure XLSX exports
OUTPUT_DIR  = "output"         # folder where HTML dashboard is written

# Tenant budget in EUR per month — set to None to disable budget alerts
TENANT_BUDGET_EUR = None       # e.g. 300.0  →  tenants over budget are highlighted red

# Friendly names for tenant GUIDs (optional — GUIDs shown if not mapped)
# Format: { "full-guid-here": "Friendly Name" }
TENANT_NAMES = {
    # "01d99f43-33d6-4182-8124-91480833455e": "Acme Corp",
    # "34076cf2-65e8-402e-9263-8b82276fdfe6": "Beta Client",
}

# Resource group name patterns that identify SHARED infrastructure.
# Any RG whose name matches one of these patterns is treated as shared
# and its cost is split equally across all tenants.
# Python regex — case-insensitive.
SHARED_RG_PATTERNS = [
    r"^prod_cx_rg$",
    r"^prod_cx_aks_nodes_rg$",
    r"^prod_cx_backup_rg$",
    r"^prod_cx_util_rg$",
    r"^rg-amba",
    r"^networkwatcher",
]

# Resources with NO resource group (NaN) are always treated as shared
# (e.g. Reserved VM Instances, Reservations — these have subscription scope)
TREAT_UNASSIGNED_AS_SHARED = True

# Pattern that identifies TENANT resource groups
# Default: any RG starting with "tenant-"
TENANT_RG_PREFIX = "tenant-"

# ============================================================
# END OF CONFIGURATION
# ============================================================

import os
import re
import sys
import json
import glob
import calendar
import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("❌  pandas not installed. Run:  pip install pandas openpyxl")


# ── Helpers ──────────────────────────────────────────────────

def rg_category(rg_name: str) -> str:
    """Return 'tenant', 'shared', or 'unassigned'."""
    if pd.isna(rg_name) or str(rg_name).strip() in ("", "nan"):
        return "unassigned" if TREAT_UNASSIGNED_AS_SHARED else "other"
    rg = str(rg_name).strip()
    if rg.lower().startswith(TENANT_RG_PREFIX.lower()):
        return "tenant"
    for pattern in SHARED_RG_PATTERNS:
        if re.match(pattern, rg, re.IGNORECASE):
            return "shared"
    return "shared"   # anything that isn't a tenant RG is treated as shared


def tenant_id_from_rg(rg_name: str) -> str:
    """Extract tenant GUID from RG name like 'tenant-GUID-rg'."""
    s = str(rg_name).strip()
    s = re.sub(r"^tenant-", "", s, flags=re.IGNORECASE)
    s = re.sub(r"-rg$", "", s, flags=re.IGNORECASE)
    return s


def tenant_label(tid: str) -> str:
    """Return friendly name if configured, else shortened GUID."""
    if tid in TENANT_NAMES:
        return TENANT_NAMES[tid]
    return tid[:8] + "…"


def fmt_month_label(month_key: str) -> str:
    """Turn '2026-04' into 'April 2026'."""
    try:
        y, m = month_key.split("-")
        return f"{calendar.month_name[int(m)]} {y}"
    except Exception:
        return month_key


# ── File detection ────────────────────────────────────────────

def detect_file(filepath: str):
    """
    Returns (file_type, month_key) where file_type is one of:
        'monthly_rg'  — monthly RG aggregation
        'daily_rg'    — daily RG aggregation
        'resource'    — per-resource detail
        None          — not a recognised Azure Cost export
    """
    try:
        summary = pd.read_excel(filepath, sheet_name="Summary", header=None)
        data    = pd.read_excel(filepath, sheet_name="Data", nrows=5)
    except Exception:
        return None, None

    cols      = [str(c).lower() for c in data.columns]
    sum_text  = " ".join(str(v) for v in summary.values.flatten() if pd.notna(v)).lower()

    # Resource-level export has ResourceId column
    if "resourceid" in cols:
        file_type = "resource"
        # Month from the summary start-date line
        match = re.search(r"(\w+)\s+(\d{1,2}),?\s+(\d{4})", sum_text)
        if match:
            try:
                dt = datetime.datetime.strptime(
                    f"{match.group(1)} {match.group(2)} {match.group(3)}", "%b %d %Y"
                )
                return file_type, f"{dt.year}-{dt.month:02d}"
            except Exception:
                pass
        return file_type, "unknown"

    # RG-grouped export has UsageDate column
    if "usagedate" in cols:
        granularity = "monthly" if "granularity: monthly" in sum_text else "daily"
        file_type   = "monthly_rg" if granularity == "monthly" else "daily_rg"
        try:
            date = pd.to_datetime(data["UsageDate"].dropna().iloc[0])
            return file_type, f"{date.year}-{date.month:02d}"
        except Exception:
            return file_type, "unknown"

    return None, None


# ── Data processing ───────────────────────────────────────────

def process_monthly_rg(df: pd.DataFrame):
    """Process the monthly RG sheet and return per-RG summary rows."""
    df = df[["ResourceGroupName", "Cost", "CostUSD"]].copy()
    df["category"] = df["ResourceGroupName"].apply(rg_category)
    return df


def process_daily_rg(df: pd.DataFrame):
    """Process the daily RG sheet."""
    df = df[["UsageDate", "ResourceGroupName", "Cost", "CostUSD"]].copy()
    df["UsageDate"]  = pd.to_datetime(df["UsageDate"]).dt.date.astype(str)
    df["category"]   = df["ResourceGroupName"].apply(rg_category)
    return df


def process_resource(df: pd.DataFrame):
    """Process the resource-level sheet."""
    wanted = ["Resource", "ResourceType", "ResourceGroupName",
              "ServiceName", "Cost", "CostUSD"]
    df = df[[c for c in wanted if c in df.columns]].copy()
    df["ResourceGroupName"] = df["ResourceGroupName"].fillna("(No RG / Unassigned)")
    df["category"] = df["ResourceGroupName"].apply(rg_category)
    return df


def build_month_payload(month_key: str,
                        monthly_df, daily_df, resource_df):
    """
    Combine the three data frames into a single month payload dict.
    All monetary values are in EUR (the 'Cost' column from Azure).
    """
    label = fmt_month_label(month_key)

    # ── Monthly RG summary ──────────────────────────────────
    if monthly_df is not None:
        mrg = process_monthly_rg(monthly_df)
    else:
        # Fall back to aggregating daily data
        if daily_df is not None:
            mrg = process_daily_rg(daily_df).groupby(
                ["ResourceGroupName", "category"], as_index=False
            )[["Cost", "CostUSD"]].sum()
        else:
            print(f"  ⚠  No monthly or daily RG data for {label} — skipping")
            return None

    tenant_rows = mrg[mrg["category"] == "tenant"]
    shared_rows = mrg[mrg["category"] != "tenant"]

    n_tenants        = len(tenant_rows)
    shared_total_eur = shared_rows["Cost"].sum()
    shared_total_usd = shared_rows["CostUSD"].sum()

    if n_tenants == 0:
        print(f"  ⚠  No tenant RGs found for {label} — skipping")
        return None

    per_tenant_eur = shared_total_eur / n_tenants
    per_tenant_usd = shared_total_usd / n_tenants

    # Tenant monthly summary
    tenant_summary = []
    for _, row in tenant_rows.iterrows():
        rg  = str(row["ResourceGroupName"])
        tid = tenant_id_from_rg(rg)
        own_eur   = round(float(row["Cost"]),   2)
        total_eur = round(own_eur + per_tenant_eur, 2)
        over_budget = (
            TENANT_BUDGET_EUR is not None and total_eur > TENANT_BUDGET_EUR
        )
        tenant_summary.append({
            "rg":           rg,
            "tenant":       tid,
            "label":        tenant_label(tid),
            "own_eur":      own_eur,
            "own_usd":      round(float(row["CostUSD"]), 2),
            "shared_eur":   round(per_tenant_eur, 2),
            "total_eur":    total_eur,
            "over_budget":  over_budget,
        })
    tenant_summary.sort(key=lambda x: x["total_eur"], reverse=True)

    # Shared RG breakdown
    shared_breakdown = []
    for _, row in shared_rows.iterrows():
        rg = str(row["ResourceGroupName"])
        if rg == "nan":
            rg = "(No RG / Unassigned)"
        shared_breakdown.append({
            "rg":             rg,
            "eur":            round(float(row["Cost"]),   2),
            "usd":            round(float(row["CostUSD"]), 2),
            "per_tenant_eur": round(float(row["Cost"]) / n_tenants, 2),
        })
    shared_breakdown.sort(key=lambda x: x["eur"], reverse=True)

    # ── Daily records ───────────────────────────────────────
    daily_records  = []
    daily_totals   = []

    if daily_df is not None:
        drg = process_daily_rg(daily_df)
        # Shared cost per day
        shared_daily = (
            drg[drg["category"] != "tenant"]
            .groupby("UsageDate")[["Cost", "CostUSD"]]
            .sum()
            .reset_index()
            .rename(columns={"Cost": "sh_eur", "CostUSD": "sh_usd"})
        )
        shared_daily["pt_eur"] = shared_daily["sh_eur"] / n_tenants

        tenant_daily = drg[drg["category"] == "tenant"].copy()
        tenant_daily = tenant_daily.merge(shared_daily, on="UsageDate", how="left")

        for _, row in tenant_daily.iterrows():
            rg  = str(row["ResourceGroupName"])
            tid = tenant_id_from_rg(rg)
            own_eur    = round(float(row["Cost"]),  4)
            shared_eur = round(float(row.get("pt_eur", 0) or 0), 4)
            daily_records.append({
                "date":       str(row["UsageDate"]),
                "tenant":     tid,
                "label":      tenant_label(tid),
                "own_eur":    own_eur,
                "shared_eur": shared_eur,
                "total_eur":  round(own_eur + shared_eur, 4),
            })

        # Aggregate daily totals across all tenants
        day_sums = {}
        for r in daily_records:
            day_sums[r["date"]] = day_sums.get(r["date"], 0.0) + r["total_eur"]
        daily_totals = [
            {"date": d, "total_eur": round(v, 2)}
            for d, v in sorted(day_sums.items())
        ]

    # ── Resource detail ─────────────────────────────────────
    resource_detail = []
    if resource_df is not None:
        res = process_resource(resource_df)
        # Only group by columns that are actually present — Azure export
        # column names vary between months and portal versions.
        group_cols = [c for c in
                      ["ResourceGroupName", "Resource", "ResourceType", "ServiceName"]
                      if c in res.columns]
        missing = [c for c in ["ResourceGroupName", "Resource", "ResourceType", "ServiceName"]
                   if c not in res.columns]
        if missing:
            print(f"      ℹ  Resource file columns not found (skipped in groupby): {missing}")
        if "Cost" not in res.columns:
            print(f"  ⚠  Resource file for {label} has no Cost column — skipping resource detail")
        else:
            grouped = (
                res.groupby(group_cols)["Cost"].sum()
                .reset_index()
                .sort_values("Cost", ascending=False)
                .head(300)
            )
            for _, row in grouped.iterrows():
                resource_detail.append({
                    "rg":       str(row.get("ResourceGroupName", "(No RG / Unassigned)")),
                    "resource": str(row.get("Resource", "-")),
                    "type":     str(row.get("ResourceType", "-")),
                    "service":  str(row.get("ServiceName", "-")),
                    "eur":      round(float(row["Cost"]), 3),
                })

    return {
        "key":                 month_key,
        "label":               label,
        "n_tenants":           n_tenants,
        "total_eur":           round(shared_total_eur + sum(t["own_eur"] for t in tenant_summary), 2),
        "shared_total_eur":    round(shared_total_eur, 2),
        "per_tenant_eur":      round(per_tenant_eur, 2),
        "tenant_own_total_eur":round(sum(t["own_eur"] for t in tenant_summary), 2),
        "budget_eur":          TENANT_BUDGET_EUR,
        "tenant_summary":      tenant_summary,
        "shared_breakdown":    shared_breakdown,
        "daily_records":       daily_records,
        "daily_totals":        daily_totals,
        "resource_detail":     resource_detail,
    }


# ── Month-over-month comparison ───────────────────────────────

def build_mom_comparison(months: list) -> list:
    result = []
    for i, m in enumerate(months):
        prev = months[i - 1] if i > 0 else None
        delta_eur     = round(m["total_eur"] - prev["total_eur"], 2) if prev else None
        delta_pct     = round((delta_eur / prev["total_eur"]) * 100, 1) if prev and prev["total_eur"] else None
        delta_tenants = m["n_tenants"] - prev["n_tenants"] if prev else None
        result.append({
            "key":            m["key"],
            "label":          m["label"],
            "total_eur":      m["total_eur"],
            "n_tenants":      m["n_tenants"],
            "shared_eur":     m["shared_total_eur"],
            "per_tenant_eur": m["per_tenant_eur"],
            "delta_eur":      delta_eur,
            "delta_pct":      delta_pct,
            "delta_tenants":  delta_tenants,
        })
    return result


# ── Main scanning logic ───────────────────────────────────────

def scan_exports(exports_dir: str) -> dict:
    """
    Scan the exports folder, detect each file's type and month,
    and return a dict keyed by month_key → {"monthly_rg": ..., "daily_rg": ..., "resource": ...}
    """
    pattern = os.path.join(exports_dir, "*.xlsx")
    files   = sorted(glob.glob(pattern))

    if not files:
        sys.exit(f"❌  No .xlsx files found in '{exports_dir}/'\n"
                 f"    Create the folder and drop your Azure Cost exports there.")

    month_files = {}  # month_key → {type: filepath}

    print(f"\n📂  Scanning '{exports_dir}/' — {len(files)} file(s) found\n")
    for fp in files:
        ftype, month_key = detect_file(fp)
        fname = os.path.basename(fp)
        if ftype is None:
            print(f"  ⚠  Skipping (not a recognised Azure export): {fname}")
            continue
        if month_key == "unknown":
            print(f"  ⚠  Could not detect month for: {fname}")
            continue

        if month_key not in month_files:
            month_files[month_key] = {}
        if ftype in month_files[month_key]:
            print(f"  ⚠  Duplicate {ftype} file for {month_key}: {fname} — using latest")
        month_files[month_key][ftype] = fp
        print(f"  ✅  [{ftype:<12}] [{month_key}]  {fname}")

    return month_files


# ── HTML dashboard generation ─────────────────────────────────

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Azure FinOps Dashboard — DSS_CX_PROD</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root{--bg:#080d18;--sur:#0f1623;--s2:#141c2e;--s3:#1a2440;--bd:#1c2a3e;--b2:#243350;
  --ac:#3b82f6;--a2:#60a5fa;--a3:#93c5fd;--gn:#10b981;--g2:#34d399;
  --or:#f59e0b;--o2:#fbbf24;--rd:#ef4444;--r2:#f87171;--pu:#8b5cf6;--p2:#a78bfa;
  --tx:#dde4f0;--t2:#8899b4;--t3:#3d5070;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--tx);font-family:var(--sans);font-size:13px;min-height:100vh}
.hdr{background:var(--sur);border-bottom:1px solid var(--bd);padding:0 28px;height:58px;
  display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200}
.hdr-l{display:flex;align-items:center;gap:14px}
.hdr-title{font-size:15px;font-weight:600;letter-spacing:-.3px}
.hdr-sub{font-size:10px;color:var(--t3);font-family:var(--mono);margin-top:2px}
.eur-badge{background:#1a3a2a;border:1px solid #0f6e56;color:#34d399;font-size:10px;
  font-family:var(--mono);padding:2px 8px;border-radius:4px;font-weight:600}
.mstrip{display:flex;gap:4px;align-items:center}
.mbtn{background:var(--s2);border:1px solid var(--bd);color:var(--t2);padding:5px 12px;
  border-radius:5px;cursor:pointer;font-size:11px;font-family:var(--sans);transition:all .15s}
.mbtn:hover{border-color:var(--b2);color:var(--tx)}
.mbtn.active{background:var(--ac);border-color:var(--ac);color:#fff;font-weight:600}
.main{padding:22px 28px;max-width:1600px}
.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:20px}
.card{background:var(--sur);border:1px solid var(--bd);border-radius:10px;padding:16px;
  position:relative;overflow:hidden}
.card::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--ca,var(--ac))}
.cl{font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:.08em;font-weight:500}
.cv{font-size:22px;font-weight:600;font-family:var(--mono);margin:6px 0 2px;letter-spacing:-1px}
.cs{font-size:10px;color:var(--t2)}
.delta{display:inline-flex;align-items:center;gap:2px;font-size:10px;font-family:var(--mono);
  padding:2px 5px;border-radius:3px;margin-left:5px;vertical-align:middle}
.delta.up{background:#2d1a1a;color:var(--rd)}.delta.dn{background:#1a2d1a;color:var(--gn)}
.banner{background:#151500;border:1px solid #3d3a00;border-radius:8px;padding:9px 14px;
  margin-bottom:16px;font-size:11px;color:#c9b84c;display:flex;align-items:center;gap:8px}
.tabs{display:flex;gap:3px;border-bottom:1px solid var(--bd);margin-bottom:20px}
.tab{padding:9px 18px;border-radius:7px 7px 0 0;cursor:pointer;font-size:12px;color:var(--t3);
  border:1px solid transparent;border-bottom:none;transition:all .15s;font-weight:500;user-select:none}
.tab:hover{color:var(--t2);background:var(--sur)}
.tab.active{color:var(--a2);background:var(--sur);border-color:var(--bd)}
.panel{display:none}.panel.active{display:block}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.cb{background:var(--sur);border:1px solid var(--bd);border-radius:10px;padding:18px}
.cb h3{font-size:11px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.07em;margin-bottom:2px}
.cb-sub{font-size:10px;color:var(--t3);margin-bottom:14px}
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:11px}
th{text-align:left;padding:8px 12px;border-bottom:1px solid var(--bd);color:var(--t3);
  font-weight:500;text-transform:uppercase;letter-spacing:.07em;font-size:10px;cursor:pointer;white-space:nowrap}
th:hover{color:var(--ac)}
td{padding:8px 12px;border-bottom:1px solid #0d1420;font-family:var(--mono);font-size:11px}
tr:hover td{background:var(--s2)}
tfoot td{border-top:2px solid var(--b2);border-bottom:none;font-weight:600}
.pill{display:inline-block;padding:2px 7px;border-radius:4px;font-size:9px;font-weight:600;text-transform:uppercase}
.ps{background:#1a2e48;color:#7dd3fc}.pt{background:#1a3028;color:#6ee7b7}
.pu{background:#2d1f3d;color:#c4b5fd}.pr{background:#2d1a1a;color:#f87171}
.filters{display:flex;gap:10px;align-items:center;margin-bottom:16px;flex-wrap:wrap}
.fl{font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:.08em}
select{background:var(--s2);border:1px solid var(--bd);color:var(--tx);padding:6px 10px;
  border-radius:5px;font-size:11px;font-family:var(--sans);outline:none;cursor:pointer}
select:focus{border-color:var(--ac)}
.inst{background:var(--s2);border:1px solid var(--b2);border-radius:8px;padding:16px;margin-bottom:16px}
.inst h4{font-size:12px;font-weight:600;color:var(--a2);margin-bottom:10px}
.is{display:flex;gap:10px;margin-bottom:8px;font-size:12px;color:var(--t2);line-height:1.5}
.in{background:var(--ac);color:#fff;width:20px;height:20px;border-radius:50%;
  display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0;margin-top:1px}
code{background:var(--s3);border:1px solid var(--bd);padding:2px 6px;border-radius:4px;
  font-family:var(--mono);font-size:10px;color:var(--a3)}
.tbar{height:5px;border-radius:2px;min-width:2px;display:inline-block}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg)}
::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px}
@media(max-width:900px){.cards{grid-template-columns:repeat(2,1fr)}.g2,.g3,.g4{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="hdr">
  <div class="hdr-l">
    <svg width="26" height="26" viewBox="0 0 96 96"><rect width="96" height="96" rx="18" fill="#0078d4"/><text x="48" y="66" font-size="52" text-anchor="middle" fill="#fff" font-family="Arial" font-weight="700">A</text></svg>
    <div><div class="hdr-title">Azure FinOps Dashboard</div><div class="hdr-sub" id="sub-text">DSS_CX_PROD</div></div>
    <span class="eur-badge">€ EUR</span>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <label style="font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:.08em">Month</label>
    <select id="mselect" onchange="sm(parseInt(this.value))" style="background:var(--s2);border:1px solid var(--b2);color:var(--tx);padding:5px 10px;border-radius:5px;font-size:12px;font-family:var(--sans);outline:none;cursor:pointer"></select>
  </div>
</div>

<div class="main">
  <div class="banner" id="banner" style="display:none">
    ⚠ Some months contain simulated data — replace with real Azure exports and re-run <code>process_costs.py</code>
  </div>
  <div class="cards" id="cards"></div>
  <div class="tabs">
    <div class="tab active"  onclick="sw('overview')">📊 Overview</div>
    <div class="tab"         onclick="sw('mom')">📈 MoM Comparison</div>
    <div class="tab"         onclick="sw('tenants')">🏢 Tenants</div>
    <div class="tab"         onclick="sw('shared')">🔗 Shared Infra</div>
    <div class="tab"         onclick="sw('daily')">📅 Daily Trend</div>
    <div class="tab"         onclick="sw('resources')">📦 Resources</div>
    <div class="tab"         onclick="sw('howto')">⚙️ How to Update</div>
  </div>

  <!-- OVERVIEW -->
  <div class="panel active" id="panel-overview">
    <div class="g2">
      <div class="cb"><h3>Total cost per tenant</h3><div class="cb-sub">Own resources + equal share of common infra (EUR)</div><div style="position:relative;height:300px"><canvas id="c-tenant-bar"></canvas></div></div>
      <div class="cb"><h3>Own vs shared stack</h3><div class="cb-sub">Blue = tenant-specific · Orange = shared allocation</div><div style="position:relative;height:300px"><canvas id="c-stacked"></canvas></div></div>
    </div>
    <div class="g2" id="overview-shared-trend-g2">
      <div class="cb"><h3>Shared RG distribution</h3><div class="cb-sub">How the shared pool breaks down by resource group</div><div style="position:relative;height:260px"><canvas id="c-sh-donut"></canvas></div></div>
      <div class="cb" id="overview-trend-card"><h3>Monthly cost trend</h3><div class="cb-sub">Total subscription cost across all loaded months</div><div style="position:relative;height:260px"><canvas id="c-trend"></canvas></div></div>
    </div>
  </div>

  <!-- MOM -->
  <div class="panel" id="panel-mom">
    <div class="g4" id="mom-cards"></div>
    <div class="g2">
      <div class="cb"><h3>Monthly spend (EUR)</h3><div class="cb-sub">Total cost per month</div><div style="position:relative;height:240px"><canvas id="c-mom-bar"></canvas></div></div>
      <div class="cb"><h3>Tenant count growth</h3><div class="cb-sub">Active tenant RGs per month</div><div style="position:relative;height:240px"><canvas id="c-growth"></canvas></div></div>
    </div>
    <div class="cb" style="margin-bottom:16px"><h3>Month-over-month table</h3><div id="mom-tbl"></div></div>
    <div class="cb" id="tenant-mom-box"><h3>Per-tenant cost across months</h3><div class="cb-sub">Tenants present in all months shown</div><div style="position:relative;height:280px"><canvas id="c-tenant-mom"></canvas></div></div>
  </div>

  <!-- TENANTS -->
  <div class="panel" id="panel-tenants">
    <div class="filters">
      <span class="fl">Sort:</span>
      <select id="t-sort" onchange="rTenants()">
        <option value="total">Total Cost</option><option value="own">Own Cost</option>
        <option value="shared">Shared Alloc</option><option value="name">Tenant ID</option>
      </select>
    </div>
    <div class="cb">
      <h3>Tenant cost breakdown — <span id="t-month-lbl"></span></h3>
      <div class="cb-sub">Own cost + <span id="t-share-lbl"></span> shared ÷ <span id="t-count-lbl"></span> tenants = <span id="t-pt-lbl"></span> per tenant</div>
      <div class="tw" id="t-tbl"></div>
    </div>
  </div>

  <!-- SHARED -->
  <div class="panel" id="panel-shared">
    <div class="g2">
      <div class="cb"><h3>Shared infrastructure (EUR)</h3><div class="cb-sub">Resources used equally by all tenants</div><div style="position:relative;height:280px"><canvas id="c-sh-bar"></canvas></div></div>
      <div class="cb"><h3>Per-tenant allocation (EUR)</h3><div class="cb-sub">Each tenant's share from each shared RG</div><div style="position:relative;height:280px"><canvas id="c-sh-alloc"></canvas></div></div>
    </div>
    <div class="cb"><h3>Shared RG detail</h3><div class="tw" id="sh-tbl"></div></div>
  </div>

  <!-- DAILY -->
  <div class="panel" id="panel-daily">
    <div class="filters">
      <span class="fl">Tenant:</span>
      <select id="d-tenant" onchange="rDaily()"><option value="all">All tenants (aggregate)</option></select>
    </div>
    <div class="cb" style="margin-bottom:16px"><h3>Daily cost trend (EUR)</h3><div class="cb-sub">Daily spend — own + shared allocation</div><div style="position:relative;height:300px"><canvas id="c-daily"></canvas></div></div>
    <div class="cb"><h3>Daily table</h3><div class="tw" id="d-tbl"></div></div>
  </div>

  <!-- RESOURCES -->
  <div class="panel" id="panel-resources">
    <div class="filters">
      <span class="fl">RG:</span><select id="r-rg" onchange="rRes()"><option value="all">All</option></select>
      <span class="fl">Service:</span><select id="r-svc" onchange="rRes()"><option value="all">All</option></select>
      <span class="fl">Min €:</span>
      <select id="r-min" onchange="rRes()">
        <option value="0">All</option><option value="1">≥ €1</option>
        <option value="10">≥ €10</option><option value="50">≥ €50</option>
      </select>
    </div>
    <div class="cb"><h3>Resource-level costs (EUR)</h3><div class="tw" id="r-tbl"></div></div>
  </div>

  <!-- HOW TO -->
  <div class="panel" id="panel-howto">
    <div class="inst">
      <h4>📥 Monthly update — 5 minutes, first-of-every-month</h4>
      <div class="is"><div class="in">1</div><div>Azure Portal → Cost Management + Billing → Cost analysis → set scope to your subscription</div></div>
      <div class="is"><div class="in">2</div><div><strong>Export 1 (Monthly RG):</strong> Group by = ResourceGroupName · Granularity = Monthly · full month date range → Download → Excel (.xlsx)</div></div>
      <div class="is"><div class="in">3</div><div><strong>Export 2 (Daily RG):</strong> Same but Granularity = Daily → Download → Excel</div></div>
      <div class="is"><div class="in">4</div><div><strong>Export 3 (Resources):</strong> Switch view to "Resources" · same date range → Download → Excel</div></div>
      <div class="is"><div class="in">5</div><div>Drop all 3 files into the <code>exports/</code> folder (any filename — auto-detected)</div></div>
      <div class="is"><div class="in">6</div><div>Run: <code>python process_costs.py</code></div></div>
      <div class="is"><div class="in">7</div><div>Open the new file in <code>output/</code> — share it by email, Teams, or SharePoint</div></div>
    </div>
    <div class="inst">
      <h4>🌐 Option A · Vercel (Recommended — free, permanent URL)</h4>
      <div class="is"><div class="in">1</div><div>Go to <strong>vercel.com</strong> → Sign Up free (use your email or GitHub)</div></div>
      <div class="is"><div class="in">2</div><div>Click <strong>Add New → Project</strong> → drag and drop your <code>output/</code> folder into the upload area</div></div>
      <div class="is"><div class="in">3</div><div>Vercel detects <code>index.html</code> automatically → click <strong>Deploy</strong></div></div>
      <div class="is"><div class="in">4</div><div>Your permanent URL is ready: <code>https://your-project.vercel.app</code> — share this with clients and managers</div></div>
      <div class="is"><div class="in">5</div><div><strong>Every month:</strong> run script → go to Vercel → your project → <strong>Deployments → Upload</strong> → drop new <code>output/</code> folder → URL stays the same</div></div>
    </div>
    <div class="inst">
      <h4>🐙 Option B · GitHub Pages (free with public repo)</h4>
      <div class="is"><div class="in">1</div><div>Create a free account at <strong>github.com</strong> → New repository → name it <code>finops-dashboard</code></div></div>
      <div class="is"><div class="in">2</div><div>Upload <code>index.html</code> from your <code>output/</code> folder to the repo root (drag and drop on the web UI)</div></div>
      <div class="is"><div class="in">3</div><div>Go to <strong>Settings → Pages → Branch: main → Save</strong></div></div>
      <div class="is"><div class="in">4</div><div>URL is live at: <code>https://yourusername.github.io/finops-dashboard/</code></div></div>
      <div class="is"><div class="in">5</div><div><strong>Every month:</strong> run script → upload new <code>index.html</code> → click <strong>Commit</strong> → URL auto-updates in ~1 minute</div></div>
      <div class="is"><div class="in">⚠</div><div>Free GitHub Pages requires a <strong>public</strong> repository. If cost data is sensitive, use Vercel instead (URL is not guessable).</div></div>
    </div>
    <div class="inst">
      <h4>📤 Other sharing options</h4>
      <div class="is"><div class="in">A</div><div><strong>Email:</strong> Attach <code>index.html</code> directly — opens in any browser offline, no login needed</div></div>
      <div class="is"><div class="in">B</div><div><strong>Teams / SharePoint:</strong> Upload to channel Files tab → share link in chat</div></div>
    </div>
    <div class="cb"><h3>Loaded data status</h3><div class="tw" id="status-tbl"></div></div>
  </div>
</div>

<script>
const DB = __INJECT_DATA__;
let mi = DB.months.length - 1;
const CI = {};
const TC = ['#3b82f6','#10b981','#8b5cf6','#f59e0b','#ef4444','#06b6d4','#22d3ee','#a78bfa','#f97316','#f87171','#0ea5e9','#34d399','#c084fc','#fb923c'];
const SC = ['#f59e0b','#3b82f6','#8b5cf6','#10b981','#ef4444','#06b6d4','#e879f9'];

const cur = () => mi === -1 ? DB.months[DB.months.length-1] : DB.months[mi];
const fe  = n  => '€' + Number(n).toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2});
const fs  = n  => '€' + Math.round(Number(n)).toLocaleString('en-GB');
const ts  = s  => s ? s.substring(0,8)+'…' : '?';

document.addEventListener('DOMContentLoaded',()=>{
  document.getElementById('sub-text').textContent =
    'DSS_CX_PROD · Generated ' + new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  if(DB.has_simulated) document.getElementById('banner').style.display='flex';
  // Default to "All Months" combined view when multiple months are loaded
  if(DB.months.length > 1){ mi = -1; }
  rStrip();
  if(mi === -1){ sw('mom'); rMom(); renderCards_combined(); }
  else { rAll(); }
});

function rStrip(){
  const sel = document.getElementById('mselect');
  const cur = sel.value;
  sel.innerHTML = DB.months.map((m,i)=>
    `<option value="${i}" ${i===mi?'selected':''}>${m.label}</option>`
  ).join('') + `<option value="-1" ${mi===-1?'selected':''}>All Months (Combined)</option>`;
}
function sm(i){
  mi = parseInt(i);
  rStrip();
  if(mi === -1){
    // Combined view: show the MoM tab automatically
    sw('mom');
    rMom();
    renderCards_combined();
  } else {
    rAll();
  }
}
function renderCards_combined(){
  const last = DB.months[DB.months.length-1];
  const first = DB.months[0];
  document.getElementById('cards').innerHTML = `
    <div class="card" style="--ca:var(--ac)"><div class="cl">Months loaded</div>
      <div class="cv" style="font-size:18px">${DB.months.length}</div>
      <div class="cs">${first.label} → ${last.label}</div></div>
    <div class="card" style="--ca:var(--or)"><div class="cl">Latest total</div>
      <div class="cv" style="font-size:18px">${fe(last.total_eur)}</div>
      <div class="cs">${last.label}</div></div>
    <div class="card" style="--ca:var(--gn)"><div class="cl">MoM change</div>
      <div class="cv" style="font-size:18px">${DB.mom[DB.mom.length-1].delta_pct!==null?(DB.mom[DB.mom.length-1].delta_pct>0?'+':'')+DB.mom[DB.mom.length-1].delta_pct+'%':'—'}</div>
      <div class="cs">vs previous month</div></div>
    <div class="card" style="--ca:var(--pu)"><div class="cl">Tenant growth</div>
      <div class="cv" style="font-size:18px">${first.n_tenants} → ${last.n_tenants}</div>
      <div class="cs">${first.label} → ${last.label}</div></div>
    <div class="card" style="--ca:var(--a2)"><div class="cl">Active tenants</div>
      <div class="cv" style="font-size:18px">${last.n_tenants}</div>
      <div class="cs">Latest month</div></div>
    <div class="card" style="--ca:var(--or)"><div class="cl">Latest per tenant</div>
      <div class="cv" style="font-size:18px">${fe(last.total_eur/last.n_tenants)}</div>
      <div class="cs">Incl. shared allocation</div></div>
  `;
}

function rAll(){
  const m = cur();
  rCards(m); rOverview(m); rMom(); rTenants(); rShared(m);
  rDailyFilters(m); rDaily(); rResFilters(m); rRes(); rHowTo();
}

function rCards(m){
  const prev = DB.months[mi-1];
  const d = prev ? m.total_eur - prev.total_eur : null;
  const dp = prev && prev.total_eur ? ((d/prev.total_eur)*100).toFixed(1) : null;
  const dt = prev ? m.n_tenants - prev.n_tenants : null;
  const dTag = (v,p,inv) => v===null?'': `<span class="delta ${(!inv&&v>0)||(inv&&v<0)?'up':'dn'}">${v>0?'▲':'▼'} ${Math.abs(p)}%</span>`;
  document.getElementById('cards').innerHTML = [
    {l:'Total ('+m.label+')',v:fe(m.total_eur),s:'All RGs incl. shared',ca:'var(--ac)',extra:dTag(d,dp,false)},
    {l:'Shared infra',v:fe(m.shared_total_eur),s:'Split across '+m.n_tenants+' tenants',ca:'var(--or)'},
    {l:'Per-tenant share',v:fe(m.per_tenant_eur),s:'for shared resources (equal allocation)',ca:'var(--gn)'},
    {l:'Tenant own total',v:fe(m.tenant_own_total_eur),s:'All tenant-specific RGs',ca:'var(--pu)'},
    {l:'Active tenants',v:String(m.n_tenants)+(dt!==null?`<span class="delta ${dt>0?'dn':'up'}" style="font-size:11px">${dt>0?'+':''}${dt}</span>`:''),s:'Tenant RGs found',ca:'var(--a2)'},
    {l:'Avg cost / tenant',v:fe(m.total_eur/m.n_tenants),s:'Incl. shared allocation',ca:'var(--or)'},
  ].map(c=>`<div class="card" style="--ca:${c.ca}"><div class="cl">${c.l}</div><div class="cv">${c.v}${c.extra||''}</div><div class="cs">${c.s}</div></div>`).join('');
}

function dc(id){ if(CI[id]){CI[id].destroy();delete CI[id];} }

function rOverview(m){
  const lbls = m.tenant_summary.map(t=>t.label||ts(t.tenant));
  dc('c-tenant-bar');
  CI['c-tenant-bar']=new Chart(document.getElementById('c-tenant-bar'),{type:'bar',
    data:{labels:lbls,datasets:[{label:'Total (EUR)',data:m.tenant_summary.map(t=>t.total_eur),
      backgroundColor:TC,borderRadius:4,borderSkipped:false}]},
    options:opts(false)});

  dc('c-stacked');
  CI['c-stacked']=new Chart(document.getElementById('c-stacked'),{type:'bar',
    data:{labels:lbls,datasets:[
      {label:'Own resources',data:m.tenant_summary.map(t=>t.own_eur),backgroundColor:'#3b82f6',borderRadius:0,borderSkipped:false},
      {label:'Shared allocation',data:m.tenant_summary.map(t=>t.shared_eur),backgroundColor:'#f59e0b',borderRadius:4,borderSkipped:'bottom'}
    ]},options:{...opts(true),scales:{x:{stacked:true,...ax()},y:{stacked:true,...ay()}}}});

  const nz = m.shared_breakdown.filter(r=>r.eur>0);
  dc('c-sh-donut');
  CI['c-sh-donut']=new Chart(document.getElementById('c-sh-donut'),{type:'doughnut',
    data:{labels:nz.map(r=>r.rg),datasets:[{data:nz.map(r=>r.eur),backgroundColor:SC,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{position:'right',labels:{color:'#8899b4',font:{size:9},boxWidth:8}},
        tooltip:{callbacks:{label:c=>c.label+': '+fe(c.raw)}}}}});

  const trendG2 = document.getElementById('overview-shared-trend-g2');
  const trendCard = document.getElementById('overview-trend-card');
  const showTrendComparison = (mi === -1 && DB.months.length > 1);
  if(!showTrendComparison){
    dc('c-trend');
    if(trendCard) trendCard.style.display='none';
    if(trendG2) trendG2.style.gridTemplateColumns='1fr';
    return;
  }

  if(trendCard) trendCard.style.display='';
  if(trendG2) trendG2.style.gridTemplateColumns='';
  dc('c-trend');
  CI['c-trend']=new Chart(document.getElementById('c-trend'),{type:'line',
    data:{labels:DB.months.map(m=>m.label),datasets:[
      {label:'Total',data:DB.months.map(m=>m.total_eur),borderColor:'#3b82f6',backgroundColor:'#3b82f615',fill:true,tension:.3,pointRadius:5,pointBackgroundColor:'#3b82f6',pointBorderColor:'#fff',pointBorderWidth:2},
      {label:'Shared',data:DB.months.map(m=>m.shared_total_eur),borderColor:'#f59e0b',backgroundColor:'transparent',borderDash:[4,2],tension:.3,pointRadius:4},
    ]},options:{...opts(false),plugins:{legend:{labels:{color:'#8899b4',font:{size:10}}},tooltip:{callbacks:{label:c=>c.dataset.label+': '+fe(c.raw)}}}}});
}

function opts(stack){return{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fe(c.raw)}}},
  scales:{x:ax(),y:ay()}};}
function ax(){return{ticks:{color:'#8899b4',font:{size:9},maxTicksLimit:14},grid:{color:'#1a2440'}};}
function ay(){return{ticks:{color:'#8899b4',callback:v=>fs(v),font:{size:10}},grid:{color:'#1a2440'}};}

function rMom(){
  // Hide the cross-month tenant chart when viewing a single month
  const tenantMomBoxOuter = document.getElementById('tenant-mom-box');
  if(tenantMomBoxOuter) tenantMomBoxOuter.style.display = (mi === -1 || DB.months.length > 1) ? '' : 'none';
  const mc = DB.mom;
  const lat = mc[mc.length-1]; const prv = mc[mc.length-2]||{};
  document.getElementById('mom-cards').innerHTML=[
    {l:'Latest total',v:fe(lat.total_eur),s:lat.label,ca:'var(--ac)'},
    {l:'MoM change',v:(lat.delta_pct!==null?(lat.delta_pct>0?'+':'')+lat.delta_pct+'%':'—'),
      s:lat.delta_eur!==null?fe(lat.delta_eur)+' vs prev':'First month',ca:lat.delta_eur>0?'var(--rd)':'var(--gn)'},
    {l:'New tenants',v:lat.delta_tenants!==null?(lat.delta_tenants>0?'+':'')+String(lat.delta_tenants):'—',s:'vs previous month',ca:'var(--gn)'},
    {l:'Avg per tenant',v:fe(lat.total_eur/lat.n_tenants),s:'Incl. shared',ca:'var(--or)'},
  ].map(c=>`<div class="card" style="--ca:${c.ca}"><div class="cl">${c.l}</div><div class="cv" style="font-size:19px">${c.v}</div><div class="cs">${c.s}</div></div>`).join('');

  const mMax = Math.max(...mc.map(m=>m.total_eur));
  dc('c-mom-bar');
  CI['c-mom-bar']=new Chart(document.getElementById('c-mom-bar'),{type:'bar',
    data:{labels:mc.map(m=>m.label),datasets:[{label:'Total (EUR)',data:mc.map(m=>m.total_eur),
      backgroundColor:TC.slice(0,mc.length),borderRadius:4,borderSkipped:false}]},
    options:{...opts(false),plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>c.dataset.label+': '+fe(c.raw)}}}}});

  dc('c-growth');
  CI['c-growth']=new Chart(document.getElementById('c-growth'),{type:mc.length<2?'bar':'line',
    data:{labels:mc.map(m=>m.label),datasets:[{label:'Tenants',data:mc.map(m=>m.n_tenants),
      borderColor:'#10b981',backgroundColor:'#10b98120',fill:true,tension:.3,
      pointRadius:6,pointBackgroundColor:'#10b981',pointBorderColor:'#fff',pointBorderWidth:2,
      borderRadius:mc.length<2?4:0}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.raw+' tenants'}}},
      scales:{x:ax(),y:{...ay(),min:0,ticks:{...ay().ticks,stepSize:1}}}}});

  let h=`<table><thead><tr><th>Month</th><th>Total (EUR)</th><th>MoM Change</th><th>Tenants</th><th>+/− Tenants</th><th>Per Tenant</th><th>Shared</th><th>Trend</th></tr></thead><tbody>`;
  mc.forEach((m,i)=>{
    const dc2=m.delta_eur>0?'color:var(--rd)':m.delta_eur<0?'color:var(--gn)':'color:var(--t2)';
    const tc2=m.delta_tenants>0?'color:var(--gn)':'color:var(--t2)';
    const bw=Math.round((m.total_eur/mMax)*140);
    h+=`<tr ${i===mi?'style="background:var(--s2)"':''}><td style="color:var(--a2);font-weight:600">${m.label}</td>
      <td>${fe(m.total_eur)}</td>
      <td style="${dc2}">${m.delta_eur!==null?(m.delta_eur>0?'+':'')+fe(m.delta_eur)+' ('+m.delta_pct+'%)':'—'}</td>
      <td>${m.n_tenants}</td>
      <td style="${tc2}">${m.delta_tenants!==null?(m.delta_tenants>0?'+':'')+m.delta_tenants:'—'}</td>
      <td style="color:var(--or)">${fe(m.per_tenant_eur)}</td>
      <td style="color:var(--t2)">${fe(m.shared_eur)}</td>
      <td><span class="tbar" style="width:${bw}px;background:${TC[i%TC.length]}"></span></td></tr>`;
  });
  document.getElementById('mom-tbl').innerHTML=h+'</tbody></table>';

  // Per-tenant trend only makes sense with 2+ months
  const tenantMomEl = document.getElementById('c-tenant-mom');
  const tenantMomBox = tenantMomEl.closest('.cb');
  if(DB.months.length < 2){
    tenantMomBox.innerHTML = '<h3>Per-tenant cost across months</h3>'
      + '<div style="padding:32px;text-align:center;color:var(--t3);font-size:12px">'
      + 'Add more monthly exports and re-run the script to see tenant trends across months.</div>';
    return;
  }
  const common = [...new Set(DB.months[0].tenant_summary.map(t=>t.tenant))]
    .filter(tid=>DB.months.every(m=>m.tenant_summary.find(t=>t.tenant===tid)))
    .slice(0,8);
  dc('c-tenant-mom');
  CI['c-tenant-mom']=new Chart(tenantMomEl,{type:'line',
    data:{labels:DB.months.map(m=>m.label),datasets:common.map((tid,i)=>({
      label:DB.months[0].tenant_summary.find(t=>t.tenant===tid)?.label||ts(tid),
      data:DB.months.map(m=>(m.tenant_summary.find(t=>t.tenant===tid)||{total_eur:null}).total_eur),
      borderColor:TC[i%TC.length],backgroundColor:'transparent',tension:.3,pointRadius:4,spanGaps:true,
    }))},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#8899b4',font:{size:9},boxWidth:8}},tooltip:{callbacks:{label:c=>c.dataset.label+': '+fe(c.raw)}}},
      scales:{x:ax(),y:ay()}}});
}

function rTenants(){
  const m=cur(); const s=document.getElementById('t-sort').value;
  let d=[...m.tenant_summary];
  if(s==='total')d.sort((a,b)=>b.total_eur-a.total_eur);
  else if(s==='own')d.sort((a,b)=>b.own_eur-a.own_eur);
  else if(s==='shared')d.sort((a,b)=>b.shared_eur-a.shared_eur);
  else d.sort((a,b)=>a.tenant.localeCompare(b.tenant));
  document.getElementById('t-month-lbl').textContent=m.label;
  document.getElementById('t-share-lbl').textContent=fe(m.shared_total_eur);
  document.getElementById('t-count-lbl').textContent=m.n_tenants;
  document.getElementById('t-pt-lbl').textContent=fe(m.per_tenant_eur);
  const mx=Math.max(...d.map(t=>t.total_eur));
  const gt=d.reduce((s,t)=>s+t.total_eur,0);
  let h=`<table><thead><tr><th>#</th><th>Tenant</th><th>Full ID</th><th>Own (EUR)</th><th>Shared (EUR)</th><th>Total (EUR)</th><th>% Share</th><th>Bar</th></tr></thead><tbody>`;
  d.forEach((t,i)=>{
    const pct=((t.total_eur/gt)*100).toFixed(1);
    const bw=Math.round((t.total_eur/mx)*150);
    const ow=Math.round((t.own_eur/t.total_eur)*bw);
    const ob=t.over_budget?'class="pr"':'';
    h+=`<tr>
      <td style="color:var(--t3)">${i+1}</td>
      <td style="color:var(--a2);font-weight:500">${t.label||ts(t.tenant)}</td>
      <td><code style="font-size:9px;background:var(--s2);padding:2px 5px;border-radius:3px">${t.tenant}</code></td>
      <td style="color:var(--ac)">${fe(t.own_eur)}</td>
      <td style="color:var(--or)">${fe(t.shared_eur)}</td>
      <td style="color:${t.over_budget?'var(--rd)':'var(--gn)'};font-weight:600">${fe(t.total_eur)}${t.over_budget?' <span class="pill pr">OVER BUDGET</span>':''}</td>
      <td style="color:var(--t2)">${pct}%</td>
      <td><div style="display:flex;height:5px;border-radius:2px;overflow:hidden;width:${bw}px">
        <div style="width:${ow}px;background:#3b82f6"></div>
        <div style="flex:1;background:#f59e0b"></div></div></td></tr>`;
  });
  const totOwn=d.reduce((s,t)=>s+t.own_eur,0);
  const totSh=d.reduce((s,t)=>s+t.shared_eur,0);
  h+=`</tbody><tfoot><tr><td colspan="3">TOTAL</td><td style="color:var(--ac)">${fe(totOwn)}</td><td style="color:var(--or)">${fe(totSh)}</td><td style="color:var(--gn)">${fe(totOwn+totSh)}</td><td colspan="2"></td></tr></tfoot></table>`;
  document.getElementById('t-tbl').innerHTML=h;
}

function rShared(m){
  const nz=m.shared_breakdown.filter(r=>r.eur>0);
  const mxE=Math.max(...m.shared_breakdown.map(r=>r.eur),1);
  dc('c-sh-bar');
  CI['c-sh-bar']=new Chart(document.getElementById('c-sh-bar'),{type:'bar',
    data:{labels:nz.map(r=>r.rg),datasets:[{label:'Cost (EUR)',data:nz.map(r=>r.eur),backgroundColor:SC,borderRadius:4,borderSkipped:false}]},
    options:{...opts(false),indexAxis:'y'}});
  dc('c-sh-alloc');
  CI['c-sh-alloc']=new Chart(document.getElementById('c-sh-alloc'),{type:'bar',
    data:{labels:nz.map(r=>r.rg),datasets:[{label:'Per tenant (EUR)',data:nz.map(r=>r.per_tenant_eur),backgroundColor:SC.map(c=>c+'99'),borderRadius:4,borderSkipped:false}]},
    options:{...opts(false),indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fe(c.raw)+' / tenant'}}}}});
  const tot=m.shared_breakdown.reduce((s,r)=>s+r.eur,0);
  let h=`<table><thead><tr><th>Resource Group</th><th>Total (EUR)</th><th>Per Tenant (EUR)</th><th>% of Pool</th><th>Bar</th></tr></thead><tbody>`;
  m.shared_breakdown.forEach(r=>{
    const isU=r.rg==='(No RG / Unassigned)';
    const bw=Math.round((r.eur/mxE)*120);
    h+=`<tr><td><span class="pill ${isU?'pu':'ps'}">${r.rg}</span></td>
      <td style="color:var(--or)">${fe(r.eur)}</td><td style="color:var(--ac)">${fe(r.per_tenant_eur)}</td>
      <td>${((r.eur/tot)*100).toFixed(1)}%</td>
      <td><span class="tbar" style="width:${bw}px;background:#f59e0b"></span></td></tr>`;
  });
  h+=`<tfoot><tr><td>TOTAL</td><td style="color:var(--or)">${fe(tot)}</td><td style="color:var(--ac)">${fe(m.per_tenant_eur)}</td><td>100%</td><td></td></tr></tfoot></table>`;
  document.getElementById('sh-tbl').innerHTML=h;
}

function rDailyFilters(m){
  const sel=document.getElementById('d-tenant');
  while(sel.options.length>1)sel.remove(1);
  [...new Set(m.daily_records.map(d=>d.tenant))].sort().forEach(tid=>{
    const o=document.createElement('option');o.value=tid;
    const lbl=m.tenant_summary.find(t=>t.tenant===tid)?.label||ts(tid);
    o.text=lbl+' ('+tid+')';sel.appendChild(o);});
}
function rDaily(){
  const m=cur(); const tf=document.getElementById('d-tenant').value;
  const isAll=tf==='all';
  const data=isAll?m.daily_totals:m.daily_records.filter(d=>d.tenant===tf).map(d=>({date:d.date,total_eur:d.total_eur}));
  const dates=[...new Set(data.map(d=>d.date))].sort();
  const dsets=isAll?[{label:'All tenants total',data:dates.map(d=>(m.daily_totals.find(x=>x.date===d)||{total_eur:0}).total_eur),borderColor:'#3b82f6',backgroundColor:'#3b82f615',fill:true,tension:.3,pointRadius:3}]
    :[{label:'Own',data:dates.map(d=>(m.daily_records.find(x=>x.date===d&&x.tenant===tf)||{own_eur:0}).own_eur),borderColor:'#3b82f6',fill:false,tension:.3,pointRadius:2},
      {label:'Shared',data:dates.map(d=>(m.daily_records.find(x=>x.date===d&&x.tenant===tf)||{shared_eur:0}).shared_eur),borderColor:'#f59e0b',fill:false,tension:.3,pointRadius:2,borderDash:[4,2]},
      {label:'Total',data:dates.map(d=>(m.daily_records.find(x=>x.date===d&&x.tenant===tf)||{total_eur:0}).total_eur),borderColor:'#10b981',fill:false,tension:.3,pointRadius:3}];
  dc('c-daily');
  CI['c-daily']=new Chart(document.getElementById('c-daily'),{type:'line',data:{labels:dates,datasets:dsets},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#8899b4',font:{size:10}}},tooltip:{callbacks:{label:c=>c.dataset.label+': '+fe(c.raw)}}},scales:{x:ax(),y:ay()}}});
  const rows=isAll?m.daily_totals.slice(-20).reverse():m.daily_records.filter(d=>d.tenant===tf).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,30);
  let h=isAll?`<table><thead><tr><th>Date</th><th>Total (EUR)</th></tr></thead><tbody>`:
    `<table><thead><tr><th>Date</th><th>Own (EUR)</th><th>Shared (EUR)</th><th>Total (EUR)</th></tr></thead><tbody>`;
  rows.forEach(r=>{
    h+=isAll?`<tr><td style="color:var(--t2)">${r.date}</td><td style="color:var(--gn)">${fe(r.total_eur)}</td></tr>`:
      `<tr><td style="color:var(--t2)">${r.date}</td><td style="color:var(--ac)">${fe(r.own_eur)}</td><td style="color:var(--or)">${fe(r.shared_eur)}</td><td style="color:var(--gn)">${fe(r.total_eur)}</td></tr>`;
  });
  document.getElementById('d-tbl').innerHTML=h+'</tbody></table>';
}

function rResFilters(m){
  const rg=document.getElementById('r-rg');const sv=document.getElementById('r-svc');
  while(rg.options.length>1)rg.remove(1);while(sv.options.length>1)sv.remove(1);
  [...new Set(m.resource_detail.map(r=>r.rg))].sort().forEach(v=>{const o=document.createElement('option');o.value=v;o.text=v;rg.appendChild(o);});
  [...new Set(m.resource_detail.map(r=>r.service).filter(Boolean))].sort().forEach(v=>{const o=document.createElement('option');o.value=v;o.text=v;sv.appendChild(o);});
}
function rRes(){
  const m=cur();
  const rg=document.getElementById('r-rg').value;
  const sv=document.getElementById('r-svc').value;
  const mn=parseFloat(document.getElementById('r-min').value);
  const d=m.resource_detail.filter(r=>(rg==='all'||r.rg===rg)&&(sv==='all'||r.service===sv)&&r.eur>=mn).slice(0,100);
  let h=`<table><thead><tr><th>Resource</th><th>Service</th><th>Type</th><th>Resource Group</th><th>Cost (EUR)</th></tr></thead><tbody>`;
  d.forEach(r=>{
    const iT=r.rg?.startsWith('tenant-');const isU=r.rg==='(No RG / Unassigned)';
    h+=`<tr><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--a2)">${r.resource||'-'}</td>
      <td style="color:var(--t2);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.service||'-'}</td>
      <td style="color:var(--t3);font-size:10px;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.type||'-'}</td>
      <td><span class="pill ${iT?'pt':isU?'pu':'ps'}" style="max-width:130px;display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:bottom">${r.rg||'-'}</span></td>
      <td style="color:${r.eur>100?'var(--rd)':r.eur>20?'var(--or)':'var(--t2)'};font-weight:${r.eur>10?'600':'400'}">${fe(r.eur)}</td></tr>`;
  });
  h+='</tbody></table>';
  if(m.resource_detail.length>100) h+=`<p style="text-align:center;padding:10px;color:var(--t3);font-size:10px">Top 100 of ${m.resource_detail.length} shown</p>`;
  document.getElementById('r-tbl').innerHTML=h;
}

function rHowTo(){
  let h=`<table><thead><tr><th>Month</th><th>Monthly RG</th><th>Daily RG</th><th>Resources</th><th>Tenants</th><th>Status</th></tr></thead><tbody>`;
  DB.months.forEach(m=>{
    h+=`<tr><td style="color:var(--a2)">${m.label}</td>
      <td>${m.has_monthly?'<span class="pill pt">✓ loaded</span>':'<span class="pill pu">missing</span>'}</td>
      <td>${m.has_daily?'<span class="pill pt">✓ loaded</span>':'<span class="pill pu">missing</span>'}</td>
      <td>${m.has_resources?'<span class="pill pt">✓ loaded</span>':'<span class="pill pu">missing</span>'}</td>
      <td>${m.n_tenants}</td>
      <td><span class="pill ${m.is_simulated?'pu':'pt'}">${m.is_simulated?'Simulated':'Real data'}</span></td></tr>`;
  });
  document.getElementById('status-tbl').innerHTML=h+'</tbody></table>';
}

function sw(n){
  const ns=['overview','mom','tenants','shared','daily','resources','howto'];
  document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',ns[i]===n));
  document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.id==='panel-'+n));
}
</script>
</body>
</html>"""


def generate_html(months_data: list) -> str:
    """Inject processed data into the HTML template and return the full HTML string."""
    mom = build_mom_comparison(months_data)

    # Flag months that have simulated data
    has_simulated = any(m.get("is_simulated", False) for m in months_data)

    payload = {
        "months":        months_data,
        "mom":           mom,
        "has_simulated": has_simulated,
    }

    json_str = json.dumps(payload, default=str, ensure_ascii=False)
    return DASHBOARD_HTML.replace("__INJECT_DATA__", json_str)


# ── Entry point ───────────────────────────────────────────────

def main():
    print("\n" + "="*58)
    print("  Azure FinOps — Monthly Cost Dashboard Generator")
    print("="*58)

    # Create output directory
    Path(OUTPUT_DIR).mkdir(exist_ok=True)
    Path(EXPORTS_DIR).mkdir(exist_ok=True)

    # Scan exports folder
    month_files = scan_exports(EXPORTS_DIR)

    if not month_files:
        sys.exit("❌  No valid Azure Cost exports found in the exports folder.")

    # Process each month
    months_data = []
    for month_key in sorted(month_files.keys()):
        files       = month_files[month_key]
        label       = fmt_month_label(month_key)
        print(f"\n⚙️   Processing {label} …")

        monthly_df  = None
        daily_df    = None
        resource_df = None

        if "monthly_rg" in files:
            monthly_df = pd.read_excel(files["monthly_rg"], sheet_name="Data")
            print(f"      monthly RG : {os.path.basename(files['monthly_rg'])}")
        if "daily_rg" in files:
            daily_df = pd.read_excel(files["daily_rg"], sheet_name="Data")
            print(f"      daily RG   : {os.path.basename(files['daily_rg'])}")
        if "resource" in files:
            resource_df = pd.read_excel(files["resource"], sheet_name="Data")
            print(f"      resources  : {os.path.basename(files['resource'])}")

        payload = build_month_payload(month_key, monthly_df, daily_df, resource_df)
        if payload is None:
            continue

        payload["has_monthly"]   = "monthly_rg" in files
        payload["has_daily"]     = "daily_rg" in files
        payload["has_resources"] = "resource" in files
        payload["is_simulated"]  = False

        months_data.append(payload)
        print(f"      ✅  {label}: {payload['n_tenants']} tenants · "
              f"Total €{payload['total_eur']:,.2f} · "
              f"Shared €{payload['shared_total_eur']:,.2f}")

    if not months_data:
        sys.exit("❌  No months could be processed. Check your export files.")

    generated = []   # list of (label, filepath) for the summary at the end

    # ── 1. Individual HTML file per month ────────────────────
    # Each file contains ONLY that month's data — ideal to share
    # with a specific client or for a focused monthly review.
    print(f"\n📄  Generating individual monthly files …")
    for month_payload in months_data:
        path = os.path.join(OUTPUT_DIR, f"dashboard_{month_payload['key']}.html")
        html = generate_html([month_payload])          # single-month list
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        kb = os.path.getsize(path) // 1024
        print(f"      ✅  {month_payload['label']:<14}  →  {path}  ({kb} KB)")
        generated.append((month_payload["label"], path, "individual"))

    # ── 2. index.html — all months, for GitHub Pages / Vercel hosting ──
    # This is the file you upload to Vercel or GitHub Pages.
    # Contains: dropdown month selector + MoM comparison + all tabs.
    print(f"\n🌐  Generating index.html (for Vercel / GitHub Pages) …")
    index_path    = os.path.join(OUTPUT_DIR, "index.html")
    combined_path = os.path.join(OUTPUT_DIR, "dashboard_combined.html")
    html = generate_html(months_data)
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    with open(combined_path, "w", encoding="utf-8") as f:
        f.write(html)
    kb = os.path.getsize(index_path) // 1024
    print(f"      ✅  index.html ({len(months_data)} months)  →  {index_path}  ({kb} KB)")
    print(f"      ✅  dashboard_combined.html  →  {combined_path}  (same file, kept for backwards compat)")
    generated.append((f"All months → index.html", index_path, "hosted"))

    # ── Summary ───────────────────────────────────────────────
    print(f"\n{'='*58}")
    print(f"  ✅  All files generated!")
    print(f"{'='*58}")
    print(f"\n  {'LABEL':<22}  {'TYPE':<12}  FILE")
    print(f"  {'-'*22}  {'-'*12}  {'-'*30}")
    for label, path, ftype in generated:
        print(f"  {label:<22}  {ftype:<12}  {os.path.basename(path)}")

    print(f"\n  ── Sharing guide ──────────────────────────────────────")
    print(f"  🌐 HOSTED URL (Recommended):")
    print(f"     Upload output/index.html to Vercel or GitHub Pages")
    print(f"     → one permanent URL to share with clients & managers")
    print(f"")
    print(f"  📧 EMAIL / TEAMS:")
    print(f"     Client (one month)   →  dashboard_YYYY-MM.html")
    print(f"     Manager (all months) →  index.html")
    print(f"{'='*58}\n")


if __name__ == "__main__":
    main()
