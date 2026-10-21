"""
Schnell Energy Equipments - Universal Smart Lights Dashboard & Telemetry Engine
================================================================================
Common dashboard generator for ALL projects (Trichy, Vellore, and any future projects).

Key Features:
1. Dynamic Project Discovery: Automatically detects all projects from .env
2. Real-Time Telemetry: Automatically queries live status map and scrapes liveview table on execution
3. Multi-Mode Dashboard Generation:
   - Unified Portfolio Dashboard (combines all projects into one executive email & multi-tab Excel)
   - Individual Project Dashboard (dedicated report for a specific corporation)
4. Comprehensive Excel Reporting:
   - Formatted NC Aging sheets with exact numerical Days in NC, conditional styling, and action recommendations
   - Zone breakdown and complete fleet status
5. Automated SMTP Email Dispatch with Excel Workbook Attachment.
"""

import os
import sys
import csv
import json
import argparse
import datetime
import smtplib
from collections import Counter, defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import requests
from playwright.sync_api import sync_playwright

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION & DYNAMIC PROJECT DISCOVERY
# ==============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Default known project metadata (fallback if not specified in .env)
DEFAULT_PROJECT_META = {
    "trichy": {
        "name": "Trichy",
        "display_name": "Tiruchirappalli (Trichy) Corporation",
        "project_id": "7",
        "username": "tycadmin",
        "password": "tyc12249"
    },
    "vellore": {
        "name": "Vellore",
        "display_name": "Vellore City Municipal Corporation",
        "project_id": "3",
        "username": "vcadmin",
        "password": "vel5555"
    }
}


def load_env():
    """Load configuration from root .env if present."""
    env_file = os.path.join(BASE_DIR, ".env")
    env_vars = {}
    if os.path.exists(env_file):
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    env_vars[key.strip()] = val.strip().strip('"').strip("'")
    return env_vars


def discover_projects(env_vars):
    """
    Dynamically discovers all configured projects from .env or defaults.
    Any project with <PROJECT>_USERNAME and <PROJECT>_PASSWORD is automatically registered!
    """
    projects = {}

    # Check for explicit ACTIVE_PROJECTS list in .env
    active_keys = []
    if "ACTIVE_PROJECTS" in env_vars:
        active_keys = [k.strip().lower() for k in env_vars["ACTIVE_PROJECTS"].split(",") if k.strip()]

    # Scan for project prefixes in .env (excluding SMTP)
    discovered_keys = set(active_keys)
    for k in env_vars.keys():
        if k.endswith("_USERNAME") and k != "SMTP_USERNAME":
            prefix = k[:-9].lower()
            discovered_keys.add(prefix)

    # If none found in .env, default to trichy & vellore
    if not discovered_keys:
        discovered_keys = set(DEFAULT_PROJECT_META.keys())

    for p_key in sorted(discovered_keys):
        prefix = p_key.upper()
        meta = DEFAULT_PROJECT_META.get(p_key, {})

        display_name = env_vars.get(f"{prefix}_NAME") or meta.get("display_name", f"{p_key.capitalize()} Smart Lights")
        project_id = env_vars.get(f"{prefix}_PROJECT_ID") or meta.get("project_id", "")
        username = env_vars.get(f"{prefix}_USERNAME") or meta.get("username", "")
        password = env_vars.get(f"{prefix}_PASSWORD") or meta.get("password", "")

        folder = os.path.join(BASE_DIR, p_key)
        os.makedirs(folder, exist_ok=True)

        projects[p_key] = {
            "key": p_key,
            "name": p_key.capitalize(),
            "display_name": display_name,
            "project_id": str(project_id),
            "username": username,
            "password": password,
            "folder": folder,
            "relay_csv": os.path.join(folder, f"{p_key}_relay_status.csv"),
            "all_nc_csv": os.path.join(folder, f"{p_key}_all_nc_panels.csv"),
            "devices_json": os.path.join(folder, f"project_{project_id}_devices.json"),
            "nc_json": os.path.join(folder, "nc_devices.json"),
        }

    return projects


# ==============================================================================
# REAL-TIME TELEMETRY SCRAPER
# ==============================================================================

def fetch_live_telemetry(project_config, portal_base="https://tnsmartlights.in"):
    """
    Connects to portal live, fetches API status map and scrapes liveview table.
    Saves and returns real-time records.
    """
    username = project_config["username"]
    password = project_config["password"]
    project_id = project_config["project_id"]
    display_name = project_config["display_name"]
    csv_path = project_config["relay_csv"]

    print(f"\n[REAL-TIME FETCH] Connecting to {portal_base} for {display_name}...")
    print(f"  Credentials: User={username} | Project ID={project_id}")

    # 1. API Status Map
    dashboard_status_map = {}
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer": f"{portal_base}/slm/"
    }

    print("  [Step 1/2] Fetching live dashboard status map across all zones/wards via API...")
    try:
        session.get(f"{portal_base}/slm/", headers=headers, timeout=20)
        session.post(
            f"{portal_base}/slm/home/",
            data={"Username": username, "password": password, "send": "Sign in"},
            headers=headers,
            allow_redirects=True,
            timeout=20
        )
        zones = session.post(f"{portal_base}/slm/home/loadzone/", data=str(project_id), headers=headers, timeout=20).json()
        for zone in zones:
            z_id = zone[0]
            wards = session.post(f"{portal_base}/slm/home/loadward/", data=f"{project_id},{z_id}", headers=headers, timeout=20).json()
            for ward in wards:
                w_id = ward[0]
                try:
                    devs = session.post(f"{portal_base}/slm/home/loadmap/", data=f"{project_id}!zone!{z_id}!{w_id}", headers=headers, timeout=15).json()
                    for d in devs:
                        if "uid" in d and "status" in d:
                            dashboard_status_map[d["uid"]] = str(d["status"])
                except Exception:
                    pass
        print(f"  [Step 1/2] Success: Retrieved API status map for {len(dashboard_status_map)} panels.")
    except Exception as e:
        print(f"  [Step 1/2] Warning during API status map fetch: {e}")

    # 2. Playwright Liveview Scraper
    print("  [Step 2/2] Launching Playwright browser to scrape liveview table (#rlclist)...")
    live_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        page.goto(f"{portal_base}/slm/", timeout=35000)
        page.fill('input[name="Username"]', username)
        page.fill('input[name="password"]', password)
        page.click('input[name="send"]')
        page.wait_for_url(f"{portal_base}/slm/home/", timeout=35000)

        page.goto(f"{portal_base}/slm/home/liveview/", timeout=35000)
        page.wait_for_selector("#rlclist tbody tr", timeout=40000)

        rows = page.locator("#rlclist tbody tr")
        total_rows = rows.count()
        print(f"  [Step 2/2] Success: Extracted {total_rows} real-time panel records from liveview server.")

        for i in range(total_rows):
            row = rows.nth(i)
            cells = row.locator("td")
            try:
                uid = cells.nth(3).inner_text().strip()
                zone = cells.nth(4).inner_text().strip()
                ward = cells.nth(5).inner_text().strip()
                location = cells.nth(6).inner_text().strip()
                last_update = cells.nth(19).inner_text().strip().split('\n')[0]
                previous_update = cells.nth(20).inner_text().strip().split('\n')[0]
                state = cells.nth(22).inner_text().strip()
                fault = cells.nth(23).inner_text().strip().split('\n')[0]

                if not state:
                    state = "UNKNOWN"

                db_status = dashboard_status_map.get(uid)
                if not db_status:
                    if state == "ON": db_status = "1"
                    elif state == "OFF": db_status = "0"
                    elif state in ["NO COMM", "NC"]: db_status = "2"
                    else: db_status = "0"

                live_records.append({
                    "uid": uid,
                    "zone": zone,
                    "ward": ward,
                    "location": location,
                    "state": state,
                    "dashboard_status": db_status,
                    "last_update": last_update,
                    "previous_update": previous_update,
                    "fault": fault,
                })
            except Exception:
                pass
        browser.close()

    if live_records:
        fieldnames = ["uid", "zone", "ward", "location", "state", "dashboard_status", "last_update", "previous_update", "fault"]
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for rec in live_records:
                writer.writerow(rec)
        print(f"  [Storage] Updated real-time CSV: {csv_path}")

    return live_records


# ==============================================================================
# DATA INGESTION & METRICS CALCULATION
# ==============================================================================

def process_project_telemetry(project_config, live=True, portal_base="https://tnsmartlights.in"):
    """
    Ingests telemetry records (live or cached) and computes full analytics.
    """
    csv_path = project_config["relay_csv"]

    if live:
        try:
            records = fetch_live_telemetry(project_config, portal_base)
        except Exception as e:
            print(f"  [Scrape Error] {e}. Falling back to saved CSV cache.")
            if not os.path.exists(csv_path):
                raise RuntimeError(f"Live fetch failed and no cached CSV exists at {csv_path}")
            with open(csv_path, mode="r", encoding="utf-8") as f:
                records = list(csv.DictReader(f))
    else:
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"Relay CSV not found at: {csv_path}")
        with open(csv_path, mode="r", encoding="utf-8") as f:
            records = list(csv.DictReader(f))

    # Auxiliary location map
    locations_map = {}
    if os.path.exists(project_config["all_nc_csv"]):
        try:
            with open(project_config["all_nc_csv"], mode="r", encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    uid = row.get("UID") or row.get("uid")
                    loc = row.get("Location") or row.get("location")
                    if uid and loc:
                        locations_map[uid] = loc
        except Exception:
            pass

    now = datetime.datetime.now()
    total_panels = len(records)
    on_count = 0
    off_count = 0
    nc_count = 0
    unknown_count = 0

    zone_stats = defaultdict(lambda: {"total": 0, "on": 0, "off": 0, "nc": 0, "unknown": 0})
    nc_panels = []

    for r in records:
        uid = r.get("uid", "").strip()
        zone = r.get("zone", "").strip() or "Unassigned"
        ward = r.get("ward", "").strip()
        state = r.get("state", "").strip()
        db_status = str(r.get("dashboard_status", "")).strip()
        last_update = r.get("last_update", "").strip()
        fault = r.get("fault", "").strip()
        location = r.get("location") or locations_map.get(uid, "")

        zone_stats[zone]["total"] += 1

        if db_status == "1":
            on_count += 1
            zone_stats[zone]["on"] += 1
        elif db_status == "0":
            off_count += 1
            zone_stats[zone]["off"] += 1
        elif db_status == "2":
            nc_count += 1
            zone_stats[zone]["nc"] += 1

            # Exact Days in NC calculation
            days_offline = -1
            time_display = "Never Communicated"
            is_valid_date = False

            if last_update and last_update not in ["null", "None", "-"] and not last_update.startswith("1970-01-01"):
                try:
                    dt = datetime.datetime.strptime(last_update, "%Y-%m-%d %H:%M:%S")
                    diff = now - dt
                    days_offline = max(0.0, diff.days + (diff.seconds / 86400.0))
                    is_valid_date = True

                    if days_offline < 1.0:
                        hrs = round(days_offline * 24.0, 1)
                        time_display = f"{hrs} hrs" if hrs >= 1.0 else f"{int(days_offline * 1440)} mins"
                    else:
                        time_display = f"{round(days_offline, 1):,} days"
                except Exception:
                    days_offline = -1
                    time_display = "Invalid Timestamp"

            if not is_valid_date or days_offline < 0:
                tier = "Never Communicated"
                action = "Check Initial Provisioning & SIM Registration"
            elif days_offline > 30.0:
                tier = "Critical (>30 Days)"
                action = "Field Tech Site Visit Required (Power/Hardware)"
            elif days_offline >= 7.0:
                tier = "High (7-30 Days)"
                action = "Inspect Feeder Phase & Local Gateway"
            else:
                tier = "Recent (<7 Days)"
                action = "Monitor Telecom GPRS / Daily Reset"

            nc_panels.append({
                "project_key": project_config["key"],
                "project_name": project_config["name"],
                "uid": uid,
                "zone": zone,
                "ward": ward,
                "location": location,
                "state": state,
                "fault": fault if fault else "-",
                "last_update": last_update if is_valid_date else "No History / 1970 Epoch",
                "days_offline": round(days_offline, 2) if is_valid_date else None,
                "days_offline_raw": days_offline,
                "time_display": time_display,
                "tier": tier,
                "action": action
            })
        else:
            unknown_count += 1
            zone_stats[zone]["unknown"] += 1

    nc_panels.sort(key=lambda x: (x["days_offline_raw"] if x["days_offline_raw"] >= 0 else -999999), reverse=True)
    tier_counts = Counter(p["tier"] for p in nc_panels)

    active_known = on_count + off_count
    on_pct = round((on_count / active_known * 100), 2) if active_known > 0 else 0.0
    off_pct = round((off_count / active_known * 100), 2) if active_known > 0 else 0.0
    nc_pct = round((nc_count / total_panels * 100), 2) if total_panels > 0 else 0.0
    comm_health_pct = round(((total_panels - nc_count) / total_panels * 100), 2) if total_panels > 0 else 0.0

    # Zone rows
    sorted_zones = sorted(zone_stats.keys(), key=lambda z: (int(z) if z.isdigit() else 999, z))
    zone_rows = []
    for z in sorted_zones:
        zd = zone_stats[z]
        ztot = zd["total"]
        znc = zd["nc"]
        zcomm = round(((ztot - znc) / ztot * 100), 1) if ztot > 0 else 0.0
        znc_pct = round((znc / ztot * 100), 1) if ztot > 0 else 0.0

        if znc_pct <= 10.0:
            badge = ("Optimal", "#059669", "#ecfdf5")
        elif znc_pct <= 25.0:
            badge = ("Moderate NC", "#d97706", "#fffbeb")
        else:
            badge = ("High Attention", "#dc2626", "#fef2f2")

        zone_rows.append({
            "zone": f"Zone {z}" if z.isdigit() else z,
            "total": ztot,
            "on": zd["on"],
            "off": zd["off"],
            "nc": znc,
            "on_pct": round((zd["on"] / ztot * 100), 1) if ztot > 0 else 0.0,
            "off_pct": round((zd["off"] / ztot * 100), 1) if ztot > 0 else 0.0,
            "nc_pct": znc_pct,
            "comm_health_pct": zcomm,
            "health_badge": badge
        })

    return {
        "project_config": project_config,
        "timestamp": now.strftime("%d-%b-%Y %I:%M %p"),
        "total_panels": total_panels,
        "on_count": on_count,
        "off_count": off_count,
        "nc_count": nc_count,
        "unknown_count": unknown_count,
        "on_pct": on_pct,
        "off_pct": off_pct,
        "nc_pct": nc_pct,
        "comm_health_pct": comm_health_pct,
        "zone_rows": zone_rows,
        "nc_panels": nc_panels,
        "tier_counts": tier_counts,
        "all_records": records
    }


# ==============================================================================
# UNIVERSAL EXCEL GENERATOR (OPENPYXL)
# ==============================================================================

def create_excel_report(data_list, output_filepath, is_portfolio=False):
    """
    Creates a styled Excel workbook.
    If is_portfolio=True, aggregates all projects into a unified multi-project workbook.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    navy_header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Segoe UI", size=14, bold=True, color="0F172A")
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
    data_font = Font(name="Segoe UI", size=10)
    bold_data_font = Font(name="Segoe UI", size=10, bold=True)

    fill_critical = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_critical = Font(name="Segoe UI", size=10, bold=True, color="991B1B")
    fill_high = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    font_high = Font(name="Segoe UI", size=10, bold=True, color="9A3412")
    fill_recent = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    font_recent = Font(name="Segoe UI", size=10, bold=True, color="854D0E")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0")
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    now_str = datetime.datetime.now().strftime("%d-%b-%Y %I:%M %p")

    # ----------------------------------------------------
    # TAB 1: Portfolio / Project Summary
    # ----------------------------------------------------
    ws_summary = wb.create_sheet(title="Portfolio_Overview" if is_portfolio else "Zone_Wise_Summary")
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary["A1"] = "SCHNELL ENERGY EQUIPMENTS • SMARTLIGHTS OPERATIONS"
    ws_summary["A1"].font = title_font
    ws_summary["A2"] = f"Executive Telemetry & Health Summary | Generated: {now_str}"
    ws_summary["A2"].font = subtitle_font
    ws_summary.row_dimensions[1].height = 24
    ws_summary.row_dimensions[2].height = 18

    if is_portfolio:
        headers_sum = [
            "Project / Corporation", "Project ID", "Total Panels", "Relay ON", "Relay OFF",
            "No Comm (NC)", "Relay ON %", "Relay OFF %", "NC %", "Comm Health %", "Attention Status"
        ]
        ws_summary.row_dimensions[4].height = 26
        for col_i, h in enumerate(headers_sum, 1):
            c = ws_summary.cell(row=4, column=col_i, value=h)
            c.fill = navy_header_fill; c.font = header_font; c.alignment = align_center; c.border = thin_border

        r_idx = 5
        tot_p = sum(d["total_panels"] for d in data_list)
        tot_on = sum(d["on_count"] for d in data_list)
        tot_off = sum(d["off_count"] for d in data_list)
        tot_nc = sum(d["nc_count"] for d in data_list)

        for d in data_list:
            ws_summary.row_dimensions[r_idx].height = 20
            c_proj = ws_summary.cell(r_idx, 1, d["project_config"]["display_name"])
            c_pid = ws_summary.cell(r_idx, 2, d["project_config"]["project_id"])
            c_tot = ws_summary.cell(r_idx, 3, d["total_panels"])
            c_on = ws_summary.cell(r_idx, 4, d["on_count"])
            c_off = ws_summary.cell(r_idx, 5, d["off_count"])
            c_nc = ws_summary.cell(r_idx, 6, d["nc_count"])
            c_onpct = ws_summary.cell(r_idx, 7, f"{d['on_pct']:.1f}%")
            c_offpct = ws_summary.cell(r_idx, 8, f"{d['off_pct']:.1f}%")
            c_ncpct = ws_summary.cell(r_idx, 9, f"{d['nc_pct']:.1f}%")
            c_health = ws_summary.cell(r_idx, 10, f"{d['comm_health_pct']:.1f}%")
            status_text = "Optimal" if d["nc_pct"] <= 10.0 else "Moderate NC" if d["nc_pct"] <= 25.0 else "High Attention"
            c_stat = ws_summary.cell(r_idx, 11, status_text)

            c_proj.alignment = align_left; c_proj.font = bold_data_font
            c_pid.alignment = align_center; c_pid.font = data_font
            for c in [c_tot, c_on, c_off, c_nc]:
                c.alignment = align_right; c.font = data_font; c.number_format = "#,##0"
            for c in [c_onpct, c_offpct, c_ncpct, c_health, c_stat]:
                c.alignment = align_center; c.font = data_font

            for col_i in range(1, 12):
                ws_summary.cell(r_idx, col_i).border = thin_border
            r_idx += 1

        # Total Portfolio Row
        ws_summary.row_dimensions[r_idx].height = 24
        t_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        ws_summary.cell(r_idx, 1, "PORTFOLIO TOTAL").alignment = align_left
        ws_summary.cell(r_idx, 2, f"{len(data_list)} Projects").alignment = align_center
        ws_summary.cell(r_idx, 3, tot_p).alignment = align_right
        ws_summary.cell(r_idx, 4, tot_on).alignment = align_right
        ws_summary.cell(r_idx, 5, tot_off).alignment = align_right
        ws_summary.cell(r_idx, 6, tot_nc).alignment = align_right
        active_k = tot_on + tot_off
        on_p = f"{(tot_on / active_k * 100):.1f}%" if active_k > 0 else "0.0%"
        off_p = f"{(tot_off / active_k * 100):.1f}%" if active_k > 0 else "0.0%"
        nc_p = f"{(tot_nc / tot_p * 100):.1f}%" if tot_p > 0 else "0.0%"
        hlth_p = f"{((tot_p - tot_nc) / tot_p * 100):.1f}%" if tot_p > 0 else "0.0%"
        ws_summary.cell(r_idx, 7, on_p).alignment = align_center
        ws_summary.cell(r_idx, 8, off_p).alignment = align_center
        ws_summary.cell(r_idx, 9, nc_p).alignment = align_center
        ws_summary.cell(r_idx, 10, hlth_p).alignment = align_center
        ws_summary.cell(r_idx, 11, "Fleet Overview").alignment = align_center

        for col_i in range(1, 12):
            c = ws_summary.cell(r_idx, col_i)
            c.font = bold_data_font
            c.fill = t_fill
            c.border = thin_border
            if col_i in [3, 4, 5, 6]: c.number_format = "#,##0"

    else:
        # Single project zone breakdown
        d = data_list[0]
        headers_zone = ["Zone", "Total Panels", "Relay ON", "Relay OFF", "Not Comm (NC)", "Relay ON %", "Relay OFF %", "NC %", "Health %", "Status"]
        ws_summary.row_dimensions[4].height = 26
        for col_i, h in enumerate(headers_zone, 1):
            c = ws_summary.cell(row=4, column=col_i, value=h)
            c.fill = navy_header_fill; c.font = header_font; c.alignment = align_center; c.border = thin_border

        r_idx = 5
        for z in d["zone_rows"]:
            ws_summary.row_dimensions[r_idx].height = 20
            ws_summary.cell(r_idx, 1, z["zone"]).alignment = align_left
            ws_summary.cell(r_idx, 2, z["total"]).alignment = align_right
            ws_summary.cell(r_idx, 3, z["on"]).alignment = align_right
            ws_summary.cell(r_idx, 4, z["off"]).alignment = align_right
            ws_summary.cell(r_idx, 5, z["nc"]).alignment = align_right
            ws_summary.cell(r_idx, 6, f"{z['on_pct']:.1f}%").alignment = align_center
            ws_summary.cell(r_idx, 7, f"{z['off_pct']:.1f}%").alignment = align_center
            ws_summary.cell(r_idx, 8, f"{z['nc_pct']:.1f}%").alignment = align_center
            ws_summary.cell(r_idx, 9, f"{z['comm_health_pct']:.1f}%").alignment = align_center
            ws_summary.cell(r_idx, 10, z["health_badge"][0]).alignment = align_center

            for col_i in range(1, 11):
                c = ws_summary.cell(r_idx, col_i)
                c.border = thin_border
                c.font = data_font
                if col_i in [2, 3, 4, 5]: c.number_format = "#,##0"
            r_idx += 1

        # Total row
        ws_summary.row_dimensions[r_idx].height = 24
        t_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        ws_summary.cell(r_idx, 1, "TOTAL (ALL ZONES)").alignment = align_left
        ws_summary.cell(r_idx, 2, d["total_panels"]).alignment = align_right
        ws_summary.cell(r_idx, 3, d["on_count"]).alignment = align_right
        ws_summary.cell(r_idx, 4, d["off_count"]).alignment = align_right
        ws_summary.cell(r_idx, 5, d["nc_count"]).alignment = align_right
        ws_summary.cell(r_idx, 6, f"{d['on_pct']:.1f}%").alignment = align_center
        ws_summary.cell(r_idx, 7, f"{d['off_pct']:.1f}%").alignment = align_center
        ws_summary.cell(r_idx, 8, f"{d['nc_pct']:.1f}%").alignment = align_center
        ws_summary.cell(r_idx, 9, f"{d['comm_health_pct']:.1f}%").alignment = align_center
        ws_summary.cell(r_idx, 10, "Fleet Overview").alignment = align_center

        for col_i in range(1, 11):
            c = ws_summary.cell(r_idx, col_i)
            c.font = bold_data_font
            c.fill = t_fill
            c.border = thin_border
            if col_i in [2, 3, 4, 5]: c.number_format = "#,##0"

    for col in ws_summary.columns:
        max_l = max(len(str(cell.value or "")) for cell in col if cell.row > 2)
        ws_summary.column_dimensions[get_column_letter(col[0].column)].width = max(max_l + 4, 13)

    # ----------------------------------------------------
    # TAB 2: NC Panels Aging Analysis (With Days in NC)
    # ----------------------------------------------------
    all_nc_panels = []
    for d in data_list:
        all_nc_panels.extend(d["nc_panels"])
    all_nc_panels.sort(key=lambda x: (x["days_offline_raw"] if x["days_offline_raw"] >= 0 else -999999), reverse=True)

    ws_nc = wb.create_sheet(title="All_NC_Panels_Aging" if is_portfolio else "NC_Panels_Aging_Analysis")
    ws_nc.views.sheetView[0].showGridLines = True

    ws_nc["A1"] = "SCHNELL ENERGY EQUIPMENTS • NOT-COMMUNICATING (NC) PANELS AGING REPORT"
    ws_nc["A1"].font = title_font
    ws_nc["A2"] = f"Total NC Units: {len(all_nc_panels):,} | Telemetry Refreshed: {now_str}"
    ws_nc["A2"].font = subtitle_font
    ws_nc.row_dimensions[1].height = 24
    ws_nc.row_dimensions[2].height = 18

    headers_nc = ["S.No"]
    if is_portfolio: headers_nc.append("Project")
    headers_nc.extend([
        "UID (Panel ID)", "Zone", "Ward", "Location / Landmark", "Relay State",
        "Fault Alarm", "Last Communication Date", "Days in NC (Days Offline)",
        "Formatted Duration", "Aging Priority Tier", "Recommended Action"
    ])

    ws_nc.row_dimensions[4].height = 28
    for col_i, h in enumerate(headers_nc, 1):
        c = ws_nc.cell(row=4, column=col_i, value=h)
        c.fill = navy_header_fill; c.font = header_font; c.alignment = align_center; c.border = thin_border

    curr_row = 5
    for idx, p in enumerate(all_nc_panels, 1):
        ws_nc.row_dimensions[curr_row].height = 20
        is_even = (curr_row % 2 == 0)
        r_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if is_even else PatternFill(fill_type=None)

        col_c = 1
        c_sno = ws_nc.cell(curr_row, col_c, idx); col_c += 1
        if is_portfolio:
            c_proj = ws_nc.cell(curr_row, col_c, p["project_name"]); col_c += 1
            c_proj.alignment = align_left; c_proj.font = bold_data_font; c_proj.border = thin_border
            if r_fill.fill_type: c_proj.fill = r_fill

        c_uid = ws_nc.cell(curr_row, col_c, p["uid"]); col_c += 1
        c_zone = ws_nc.cell(curr_row, col_c, f"Zone {p['zone']}" if p['zone'].isdigit() else p['zone']); col_c += 1
        c_ward = ws_nc.cell(curr_row, col_c, p["ward"]); col_c += 1
        c_loc = ws_nc.cell(curr_row, col_c, p["location"] if p["location"] else "-"); col_c += 1
        c_state = ws_nc.cell(curr_row, col_c, p["state"]); col_c += 1
        c_fault = ws_nc.cell(curr_row, col_c, p["fault"]); col_c += 1
        c_date = ws_nc.cell(curr_row, col_c, p["last_update"]); col_c += 1

        c_days = ws_nc.cell(curr_row, col_c, p["days_offline"] if p["days_offline"] is not None else "No Data"); col_c += 1
        if isinstance(p["days_offline"], (int, float)):
            c_days.number_format = "#,##0.00"

        c_dur = ws_nc.cell(curr_row, col_c, p["time_display"]); col_c += 1
        c_tier = ws_nc.cell(curr_row, col_c, p["tier"]); col_c += 1
        c_act = ws_nc.cell(curr_row, col_c, p["action"]); col_c += 1

        for c in [c_sno, c_zone, c_ward, c_state, c_fault, c_date, c_dur, c_tier]:
            c.alignment = align_center; c.font = data_font; c.border = thin_border
            if r_fill.fill_type: c.fill = r_fill

        c_uid.alignment = align_left; c_uid.font = bold_data_font; c_uid.border = thin_border
        c_loc.alignment = align_left; c_loc.font = data_font; c_loc.border = thin_border
        c_days.alignment = align_right; c_days.font = bold_data_font; c_days.border = thin_border
        c_act.alignment = align_left; c_act.font = data_font; c_act.border = thin_border
        if r_fill.fill_type:
            c_uid.fill = r_fill; c_loc.fill = r_fill; c_days.fill = r_fill; c_act.fill = r_fill

        if p["tier"] == "Critical (>30 Days)":
            c_tier.fill = fill_critical; c_tier.font = font_critical; c_days.font = font_critical
        elif p["tier"] == "High (7-30 Days)":
            c_tier.fill = fill_high; c_tier.font = font_high; c_days.font = font_high
        elif p["tier"] == "Recent (<7 Days)":
            c_tier.fill = fill_recent; c_tier.font = font_recent

        curr_row += 1

    last_col_letter = get_column_letter(len(headers_nc))
    ws_nc.auto_filter.ref = f"A4:{last_col_letter}{curr_row - 1}"
    ws_nc.freeze_panes = "A5"

    for col in ws_nc.columns:
        max_l = max(len(str(cell.value or "")) for cell in col if cell.row > 2)
        ws_nc.column_dimensions[get_column_letter(col[0].column)].width = max(max_l + 4, 12)

    # ----------------------------------------------------
    # TAB 3+: Individual Project NC & Fleet tabs if portfolio
    # ----------------------------------------------------
    if is_portfolio:
        for d in data_list:
            p_name = d["project_config"]["name"]
            ws_p = wb.create_sheet(title=f"{p_name}_NC_Panels")
            ws_p.views.sheetView[0].showGridLines = True
            ws_p["A1"] = f"{d['project_config']['display_name'].upper()} • NC PANELS"
            ws_p["A1"].font = title_font
            ws_p.row_dimensions[1].height = 22

            headers_p = ["S.No", "UID", "Zone", "Ward", "Location", "State", "Fault", "Last Update", "Days in NC", "Duration", "Tier", "Action"]
            ws_p.row_dimensions[3].height = 24
            for col_i, h in enumerate(headers_p, 1):
                c = ws_p.cell(3, col_i, h)
                c.fill = navy_header_fill; c.font = header_font; c.alignment = align_center; c.border = thin_border

            p_row = 4
            for idx, p in enumerate(d["nc_panels"], 1):
                ws_p.row_dimensions[p_row].height = 18
                ws_p.cell(p_row, 1, idx).alignment = align_center
                ws_p.cell(p_row, 2, p["uid"]).alignment = align_left
                ws_p.cell(p_row, 3, p["zone"]).alignment = align_center
                ws_p.cell(p_row, 4, p["ward"]).alignment = align_center
                ws_p.cell(p_row, 5, p["location"] if p["location"] else "-").alignment = align_left
                ws_p.cell(p_row, 6, p["state"]).alignment = align_center
                ws_p.cell(p_row, 7, p["fault"]).alignment = align_center
                ws_p.cell(p_row, 8, p["last_update"]).alignment = align_center
                
                c_d = ws_p.cell(p_row, 9, p["days_offline"] if p["days_offline"] is not None else "No Data")
                c_d.alignment = align_right
                if isinstance(p["days_offline"], (int, float)): c_d.number_format = "#,##0.00"

                ws_p.cell(p_row, 10, p["time_display"]).alignment = align_center
                ws_p.cell(p_row, 11, p["tier"]).alignment = align_center
                ws_p.cell(p_row, 12, p["action"]).alignment = align_left

                for col_i in range(1, 13):
                    c = ws_p.cell(p_row, col_i)
                    c.border = thin_border
                    c.font = data_font
                p_row += 1

            ws_p.auto_filter.ref = f"A3:L{p_row - 1}"
            ws_p.freeze_panes = "A4"
            for col in ws_p.columns:
                max_l = max(len(str(cell.value or "")) for cell in col[:100])
                ws_p.column_dimensions[get_column_letter(col[0].column)].width = max(max_l + 3, 11)

    os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
    wb.save(output_filepath)
    print(f"Generated Excel Workbook: {output_filepath}")
    return output_filepath


# ==============================================================================
# UNIVERSAL HTML EMAIL GENERATOR
# ==============================================================================

def generate_html_email(data_list, excel_filename, is_portfolio=False):
    """
    Renders an email-client compliant, responsive HTML dashboard.
    Works for both single-project and unified multi-project portfolio reports.
    """
    now_str = datetime.datetime.now().strftime("%d-%b-%Y %I:%M %p")

    # Aggregates
    tot_panels = sum(d["total_panels"] for d in data_list)
    tot_on = sum(d["on_count"] for d in data_list)
    tot_off = sum(d["off_count"] for d in data_list)
    tot_nc = sum(d["nc_count"] for d in data_list)
    active_k = tot_on + tot_off
    on_pct = round((tot_on / active_k * 100), 2) if active_k > 0 else 0.0
    off_pct = round((tot_off / active_k * 100), 2) if active_k > 0 else 0.0
    nc_pct = round((tot_nc / tot_panels * 100), 2) if tot_panels > 0 else 0.0
    health_pct = round(((tot_panels - tot_nc) / tot_panels * 100), 2) if tot_panels > 0 else 0.0

    # Aging aggregation
    comb_tier_counts = Counter()
    for d in data_list:
        comb_tier_counts.update(d["tier_counts"])

    c_crit = comb_tier_counts.get("Critical (>30 Days)", 0)
    c_high = comb_tier_counts.get("High (7-30 Days)", 0)
    c_rec = comb_tier_counts.get("Recent (<7 Days)", 0)
    c_never = comb_tier_counts.get("Never Communicated", 0)

    # Header Title
    if is_portfolio:
        project_title = f"Multi-Project Operations Portfolio ({len(data_list)} Corporations)"
        subtitle_desc = "Tiruchirappalli (Trichy) & Vellore Municipal Corporations Fleet Telemetry"
    else:
        project_title = data_list[0]["project_config"]["display_name"]
        subtitle_desc = f"Project ID: {data_list[0]['project_config']['project_id']} • Executive Fleet Relay & NC Telemetry"

    # Multi-project comparison table (if portfolio)
    portfolio_table_html = ""
    if is_portfolio:
        proj_rows_html = ""
        for d in data_list:
            st_text = "Optimal" if d["nc_pct"] <= 10.0 else "Moderate NC" if d["nc_pct"] <= 25.0 else "High Attention"
            st_color = "#059669" if d["nc_pct"] <= 10.0 else "#d97706" if d["nc_pct"] <= 25.0 else "#dc2626"
            st_bg = "#ecfdf5" if d["nc_pct"] <= 10.0 else "#fffbeb" if d["nc_pct"] <= 25.0 else "#fef2f2"

            proj_rows_html += f"""
            <tr style="border-bottom: 1px solid #e2e8f0;">
              <td style="padding: 10px 14px; font-weight: 700; color: #1e293b; font-size: 13px;">{d['project_config']['display_name']}</td>
              <td style="padding: 10px 14px; text-align: center; font-size: 12px; color: #64748b; font-family: monospace;">ID: {d['project_config']['project_id']}</td>
              <td style="padding: 10px 14px; text-align: right; font-weight: 600; color: #334155; font-size: 13px;">{d['total_panels']:,}</td>
              <td style="padding: 10px 14px; text-align: right; color: #059669; font-weight: 700; font-size: 13px;">{d['on_count']:,} <span style="font-size: 11px; font-weight: 400; color: #64748b;">({d['on_pct']}%)</span></td>
              <td style="padding: 10px 14px; text-align: right; color: #2563eb; font-weight: 700; font-size: 13px;">{d['off_count']:,} <span style="font-size: 11px; font-weight: 400; color: #64748b;">({d['off_pct']}%)</span></td>
              <td style="padding: 10px 14px; text-align: right; color: #dc2626; font-weight: 700; font-size: 13px;">{d['nc_count']:,} <span style="font-size: 11px; font-weight: 400; color: #64748b;">({d['nc_pct']}%)</span></td>
              <td style="padding: 10px 14px; text-align: center; font-weight: 700; color: #0f172a; font-size: 13px;">{d['comm_health_pct']}%</td>
              <td style="padding: 10px 14px; text-align: center;">
                <span style="display: inline-block; padding: 3px 8px; border-radius: 9999px; font-size: 11px; font-weight: 700; color: {st_color}; background-color: {st_bg}; border: 1px solid {st_color}33;">
                  {st_text}
                </span>
              </td>
            </tr>
            """

        portfolio_table_html = f"""
        <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 24px; border: 1px solid #cbd5e1; border-radius: 8px; overflow: hidden; border-collapse: separate; border-spacing: 0;">
          <thead>
            <tr style="background-color: #0f172a;">
              <th colspan="8" style="padding: 12px 14px; text-align: left; font-size: 13px; font-weight: 800; color: #ffffff; letter-spacing: 0.5px; border-bottom: 2px solid #38bdf8;">
                PROJECT-WISE COMPARISON & FLEET BREAKDOWN
              </th>
            </tr>
            <tr style="background-color: #1e293b; color: #f8fafc; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;">
              <th style="padding: 9px 14px; text-align: left;">Project / City</th>
              <th style="padding: 9px 14px; text-align: center;">ID</th>
              <th style="padding: 9px 14px; text-align: right;">Total Panels</th>
              <th style="padding: 9px 14px; text-align: right;">Relay ON</th>
              <th style="padding: 9px 14px; text-align: right;">Relay OFF</th>
              <th style="padding: 9px 14px; text-align: right;">No Comm (NC)</th>
              <th style="padding: 9px 14px; text-align: center;">Health Rate</th>
              <th style="padding: 9px 14px; text-align: center;">Status</th>
            </tr>
          </thead>
          <tbody>
            {proj_rows_html}
            <tr style="background-color: #f8fafc; border-top: 2px solid #0f172a;">
              <td colspan="2" style="padding: 12px 14px; font-weight: 800; color: #0f172a; font-size: 13px;">PORTFOLIO TOTAL</td>
              <td style="padding: 12px 14px; text-align: right; font-weight: 800; color: #0f172a; font-size: 13px;">{tot_panels:,}</td>
              <td style="padding: 12px 14px; text-align: right; font-weight: 800; color: #059669; font-size: 13px;">{tot_on:,} ({on_pct}%)</td>
              <td style="padding: 12px 14px; text-align: right; font-weight: 800; color: #2563eb; font-size: 13px;">{tot_off:,} ({off_pct}%)</td>
              <td style="padding: 12px 14px; text-align: right; font-weight: 800; color: #dc2626; font-size: 13px;">{tot_nc:,} ({nc_pct}%)</td>
              <td style="padding: 12px 14px; text-align: center; font-weight: 800; color: #0f172a; font-size: 13px;">{health_pct}%</td>
              <td style="padding: 12px 14px; text-align: center; font-size: 12px; font-weight: 700; color: #475569;">Fleet Overview</td>
            </tr>
          </tbody>
        </table>
        """

    # Zone Breakdown Tables (one per project)
    zone_tables_html = ""
    for d in data_list:
        z_rows_html = ""
        for z in d["zone_rows"]:
            b_label, b_color, b_bg = z["health_badge"]
            z_rows_html += f"""
            <tr style="border-bottom: 1px solid #e2e8f0;">
              <td style="padding: 9px 14px; font-weight: 700; color: #1e293b; font-size: 12px;">{z['zone']}</td>
              <td style="padding: 9px 14px; text-align: right; font-weight: 600; color: #334155; font-size: 12px;">{z['total']:,}</td>
              <td style="padding: 9px 14px; text-align: right; color: #059669; font-weight: 700; font-size: 12px;">{z['on']:,} <span style="font-size: 10px; font-weight: 400; color: #64748b;">({z['on_pct']}%)</span></td>
              <td style="padding: 9px 14px; text-align: right; color: #2563eb; font-weight: 700; font-size: 12px;">{z['off']:,} <span style="font-size: 10px; font-weight: 400; color: #64748b;">({z['off_pct']}%)</span></td>
              <td style="padding: 9px 14px; text-align: right; color: #dc2626; font-weight: 700; font-size: 12px;">{z['nc']:,} <span style="font-size: 10px; font-weight: 400; color: #64748b;">({z['nc_pct']}%)</span></td>
              <td style="padding: 9px 14px; text-align: center; font-weight: 700; color: #0f172a; font-size: 12px;">{z['comm_health_pct']}%</td>
              <td style="padding: 9px 14px; text-align: center;">
                <span style="display: inline-block; padding: 2px 7px; border-radius: 9999px; font-size: 10px; font-weight: 700; color: {b_color}; background-color: {b_bg}; border: 1px solid {b_color}33;">
                  {b_label}
                </span>
              </td>
            </tr>
            """

        zone_tables_html += f"""
        <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 20px; border: 1px solid #cbd5e1; border-radius: 8px; overflow: hidden; border-collapse: separate; border-spacing: 0;">
          <thead>
            <tr style="background-color: #0f172a;">
              <th colspan="7" style="padding: 10px 14px; text-align: left; font-size: 12px; font-weight: 800; color: #ffffff; letter-spacing: 0.5px; border-bottom: 2px solid #38bdf8;">
                {d['project_config']['display_name'].upper()} • ZONE-WISE BREAKDOWN
              </th>
            </tr>
            <tr style="background-color: #1e293b; color: #f8fafc; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;">
              <th style="padding: 8px 14px; text-align: left;">Zone</th>
              <th style="padding: 8px 14px; text-align: right;">Total Panels</th>
              <th style="padding: 8px 14px; text-align: right;">Relay ON</th>
              <th style="padding: 8px 14px; text-align: right;">Relay OFF</th>
              <th style="padding: 8px 14px; text-align: right;">No Comm (NC)</th>
              <th style="padding: 8px 14px; text-align: center;">Health Rate</th>
              <th style="padding: 8px 14px; text-align: center;">Status</th>
            </tr>
          </thead>
          <tbody>
            {z_rows_html}
            <tr style="background-color: #f8fafc; border-top: 2px solid #0f172a;">
              <td style="padding: 10px 14px; font-weight: 800; color: #0f172a; font-size: 12px;">SUBTOTAL ({d['project_config']['name'].upper()})</td>
              <td style="padding: 10px 14px; text-align: right; font-weight: 800; color: #0f172a; font-size: 12px;">{d['total_panels']:,}</td>
              <td style="padding: 10px 14px; text-align: right; font-weight: 800; color: #059669; font-size: 12px;">{d['on_count']:,} ({d['on_pct']}%)</td>
              <td style="padding: 10px 14px; text-align: right; font-weight: 800; color: #2563eb; font-size: 12px;">{d['off_count']:,} ({d['off_pct']}%)</td>
              <td style="padding: 10px 14px; text-align: right; font-weight: 800; color: #dc2626; font-size: 12px;">{d['nc_count']:,} ({d['nc_pct']}%)</td>
              <td style="padding: 10px 14px; text-align: center; font-weight: 800; color: #0f172a; font-size: 12px;">{d['comm_health_pct']}%</td>
              <td style="padding: 10px 14px; text-align: center; font-size: 11px; font-weight: 700; color: #475569;">City Total</td>
            </tr>
          </tbody>
        </table>
        """

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Schnell Energy Equipments • {project_title} Mail Dashboard</title>
</head>
<body style="margin: 0; padding: 0; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b;">
  <center style="width: 100%; background-color: #f1f5f9; padding: 24px 10px;">
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 780px; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 16px rgba(15, 23, 42, 0.08); border: 1px solid #e2e8f0;">
      
      <!-- HEADER -->
      <tr>
        <td style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 26px 28px; border-bottom: 3px solid #38bdf8;">
          <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td valign="middle" align="left">
                <div style="font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; color: #38bdf8; font-weight: 800; margin-bottom: 4px;">
                  SCHNELL ENERGY EQUIPMENTS • SMARTLIGHTS OPERATIONS
                </div>
                <h1 style="margin: 0; font-size: 22px; font-weight: 800; color: #ffffff; line-height: 1.25;">
                  {project_title}
                </h1>
                <div style="font-size: 13px; color: #94a3b8; margin-top: 4px;">
                  {subtitle_desc}
                </div>
              </td>
              <td valign="middle" align="right" style="width: 32%;">
                <div style="background-color: rgba(255, 255, 255, 0.08); border: 1px solid rgba(255, 255, 255, 0.16); border-radius: 8px; padding: 10px 14px; text-align: right;">
                  <div style="font-size: 10px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Executed At</div>
                  <div style="font-size: 13px; font-weight: 700; color: #38bdf8; margin-top: 2px;">{now_str}</div>
                  <div style="font-size: 11px; color: #cbd5e1; margin-top: 2px;">Scope: {len(data_list)} Active Projects</div>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- BODY CONTENT -->
      <tr>
        <td style="padding: 24px 24px 16px 24px;">

          <!-- SECTION 1: TOP KPI CARDS (Total, NC, Relay ON, Relay OFF) -->
          <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 22px;">
            <tr>
              <td colspan="4" style="padding-bottom: 10px;">
                <h2 style="margin: 0; font-size: 15px; font-weight: 800; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px;">
                  Real-Time Fleet Status
                </h2>
              </td>
            </tr>
            <tr>
              <!-- TOTAL -->
              <td width="25%" style="padding: 4px;">
                <div style="background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%); border: 1px solid #cbd5e1; border-top: 4px solid #0f172a; border-radius: 8px; padding: 14px 12px; text-align: center;">
                  <div style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: #64748b; letter-spacing: 0.5px;">TOTAL PANELS</div>
                  <div style="font-size: 26px; font-weight: 800; color: #0f172a; margin-top: 4px; line-height: 1;">{tot_panels:,}</div>
                  <div style="font-size: 11px; color: #475569; font-weight: 600; margin-top: 6px;">100% Monitored</div>
                </div>
              </td>

              <!-- RELAY ON -->
              <td width="25%" style="padding: 4px;">
                <div style="background: linear-gradient(180deg, #ffffff 0%, #f0fdf4 100%); border: 1px solid #bbf7d0; border-top: 4px solid #059669; border-radius: 8px; padding: 14px 12px; text-align: center;">
                  <div style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: #059669; letter-spacing: 0.5px;">RELAY ON</div>
                  <div style="font-size: 26px; font-weight: 800; color: #059669; margin-top: 4px; line-height: 1;">{tot_on:,}</div>
                  <div style="font-size: 11px; color: #166534; font-weight: 600; margin-top: 6px;">{on_pct}% of Active</div>
                </div>
              </td>

              <!-- RELAY OFF -->
              <td width="25%" style="padding: 4px;">
                <div style="background: linear-gradient(180deg, #ffffff 0%, #eff6ff 100%); border: 1px solid #bfdbfe; border-top: 4px solid #2563eb; border-radius: 8px; padding: 14px 12px; text-align: center;">
                  <div style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: #2563eb; letter-spacing: 0.5px;">RELAY OFF</div>
                  <div style="font-size: 26px; font-weight: 800; color: #2563eb; margin-top: 4px; line-height: 1;">{tot_off:,}</div>
                  <div style="font-size: 11px; color: #1e40af; font-weight: 600; margin-top: 6px;">{off_pct}% of Active</div>
                </div>
              </td>

              <!-- NOT COMM (NC) -->
              <td width="25%" style="padding: 4px;">
                <div style="background: linear-gradient(180deg, #ffffff 0%, #fef2f2 100%); border: 1px solid #fecaca; border-top: 4px solid #dc2626; border-radius: 8px; padding: 14px 12px; text-align: center;">
                  <div style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: #dc2626; letter-spacing: 0.5px;">NO COMM (NC)</div>
                  <div style="font-size: 26px; font-weight: 800; color: #dc2626; margin-top: 4px; line-height: 1;">{tot_nc:,}</div>
                  <div style="font-size: 11px; color: #991b1b; font-weight: 700; margin-top: 6px;">{nc_pct}% Deficit</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- SECTION 2A: MULTI-PROJECT COMPARISON TABLE (IF PORTFOLIO) -->
          {portfolio_table_html}

          <!-- SECTION 2B: ZONE-WISE BREAKDOWN TABLES -->
          {zone_tables_html}

          <!-- SECTION 3: NC AGING DISTRIBUTION -->
          <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 18px;">
            <tr>
              <td colspan="4" style="padding-bottom: 8px;">
                <h2 style="margin: 0; font-size: 15px; font-weight: 800; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px;">
                  NC Aging Analysis (Duration Panels Have Been Offline)
                </h2>
                <p style="margin: 4px 0 0 0; font-size: 12px; color: #64748b;">
                  Non-communicating units categorized by consecutive days offline as of execution:
                </p>
              </td>
            </tr>
            <tr>
              <td width="25%" style="padding: 3px;">
                <div style="background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 10px; text-align: center;">
                  <div style="font-size: 11px; font-weight: 700; color: #991b1b;">&gt; 30 DAYS (CRITICAL)</div>
                  <div style="font-size: 20px; font-weight: 800; color: #991b1b; margin-top: 2px;">{c_crit:,}</div>
                  <div style="font-size: 10px; color: #b91c1c; margin-top: 2px;">Power / HW Fault</div>
                </div>
              </td>
              <td width="25%" style="padding: 3px;">
                <div style="background-color: #fff7ed; border: 1px solid #fed7aa; border-radius: 6px; padding: 10px; text-align: center;">
                  <div style="font-size: 11px; font-weight: 700; color: #9a3412;">7 – 30 DAYS (HIGH)</div>
                  <div style="font-size: 20px; font-weight: 800; color: #9a3412; margin-top: 2px;">{c_high:,}</div>
                  <div style="font-size: 10px; color: #c2410c; margin-top: 2px;">Feeder Phase Check</div>
                </div>
              </td>
              <td width="25%" style="padding: 3px;">
                <div style="background-color: #fefce8; border: 1px solid #fef08a; border-radius: 6px; padding: 10px; text-align: center;">
                  <div style="font-size: 11px; font-weight: 700; color: #854d0e;">&lt; 7 DAYS (RECENT)</div>
                  <div style="font-size: 20px; font-weight: 800; color: #854d0e; margin-top: 2px;">{c_rec:,}</div>
                  <div style="font-size: 10px; color: #a16207; margin-top: 2px;">GPRS Glitch / Transient</div>
                </div>
              </td>
              <td width="25%" style="padding: 3px;">
                <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px; text-align: center;">
                  <div style="font-size: 11px; font-weight: 700; color: #475569;">NEVER COMMUNICATED</div>
                  <div style="font-size: 20px; font-weight: 800; color: #475569; margin-top: 2px;">{c_never:,}</div>
                  <div style="font-size: 10px; color: #64748b; margin-top: 2px;">SIM / Commissioning</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- SECTION 4: EXCEL ATTACHMENT CALLOUT BANNER -->
          <div style="background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-radius: 8px; padding: 18px 22px; margin-bottom: 12px; border-left: 6px solid #10b981;">
            <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td valign="middle" align="left">
                  <div style="font-size: 11px; text-transform: uppercase; font-weight: 800; color: #34d399; letter-spacing: 1px;">
                    ATTACHED EXCEL WORKBOOK REPORT
                  </div>
                  <div style="font-size: 16px; font-weight: 800; color: #ffffff; margin-top: 2px;">
                    {excel_filename}
                  </div>
                  <div style="font-size: 12px; color: #cbd5e1; margin-top: 4px;">
                    Contains complete inventory of all <strong>{tot_nc:,} NC panels</strong> with exact <strong>Days in NC</strong>, timestamps, ward mappings, coordinates, and diagnostic action recommendations.
                  </div>
                </td>
                <td valign="middle" align="right" style="width: 25%;">
                  <div style="background-color: rgba(16, 185, 129, 0.15); border: 1px solid #10b981; border-radius: 6px; padding: 10px 14px; text-align: center; color: #34d399; font-size: 12px; font-weight: 700;">
                    📊 XLSX Attached
                  </div>
                </td>
              </tr>
            </table>
          </div>

        </td>
      </tr>

      <!-- FOOTER -->
      <tr>
        <td style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 18px 24px; text-align: center;">
          <div style="font-size: 12px; font-weight: 700; color: #0f172a;">
            Schnell Energy Equipments • Smartlights Operations
          </div>
          <div style="font-size: 11px; color: #64748b; margin-top: 4px;">
            Confidential • Automated Mail Telemetry Dispatch • Tamil Nadu Smart Lights (tnsmartlights.in)
          </div>
        </td>
      </tr>

    </table>
  </center>
</body>
</html>
"""
    return html_content


# ==============================================================================
# SMTP EMAIL DISPATCHER
# ==============================================================================

def send_dashboard_email(html_content, excel_path, project_title, custom_recipients=None, project_key=None):
    """
    Sends the generated HTML dashboard with the attached Excel file via SMTP.
    """
    env_vars = load_env()
    smtp_server = env_vars.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(env_vars.get("SMTP_PORT", 587))
    smtp_user = env_vars.get("SMTP_USERNAME", "")
    smtp_pass = env_vars.get("SMTP_PASSWORD", "")
    sender_email = env_vars.get("SENDER_EMAIL", smtp_user)

    proj_recipients = env_vars.get(f"{project_key.upper()}_RECIPIENTS", "") if project_key else ""
    recipients_str = custom_recipients or proj_recipients or env_vars.get("RECIPIENTS", "")
    recipients = [r.strip() for r in recipients_str.split(",") if r.strip()]

    subject = f"Schnell IoT Smart Lights • {project_title} Mail Dashboard & NC Analysis - {datetime.datetime.now().strftime('%d-%b-%Y')}"

    if not smtp_user or not smtp_pass or not recipients:
        print("\n=== DRY RUN / EMAIL CONFIGURATION NOTICE ===")
        print(f"Subject       : {subject}")
        print(f"From          : {sender_email or 'Not configured'}")
        print(f"To            : {', '.join(recipients) if recipients else 'No recipients set in .env'}")
        print(f"Attachment    : {excel_path} ({os.path.getsize(excel_path):,} bytes)")
        print(f"HTML Size     : {len(html_content):,} bytes")
        print("\nTo send real emails, verify your .env file with SMTP_USERNAME, SMTP_PASSWORD, and RECIPIENTS.")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipients)

    msg_alternative = MIMEMultipart("alternative")
    msg_alternative.attach(MIMEText(html_content, "html", "utf-8"))
    msg.attach(msg_alternative)

    excel_filename = os.path.basename(excel_path)
    with open(excel_path, "rb") as f:
        att = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        att.set_payload(f.read())
        encoders.encode_base64(att)
        att.add_header("Content-Disposition", f"attachment; filename={excel_filename}")
        msg.attach(att)

    try:
        print(f"\nConnecting to SMTP server {smtp_server}:{smtp_port}...")
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=25)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=25)
            server.ehlo(); server.starttls(); server.ehlo()

        server.login(smtp_user, smtp_pass)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        print(f"Successfully dispatched Mail Dashboard to: {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"Failed to send email via SMTP: {e}")
        return False


# ==============================================================================
# MAIN DRIVER FUNCTION
# ==============================================================================

def run_common_generator(project=None, send_email=False, custom_recipients=None, live=True):
    """
    Main universal generator function:
    - If project is None, 'all', or 'portfolio': Generates Unified Multi-Project Dashboard
    - If project is specific (e.g., 'trichy', 'vellore'): Generates dedicated project report
    """
    env_vars = load_env()
    all_projects = discover_projects(env_vars)
    portal_base = env_vars.get("PORTAL_BASE_URL", "https://tnsmartlights.in")

    target_key = (project.lower() if project else "all").strip()

    # Determine execution scope
    is_portfolio = (target_key in ["all", "both", "portfolio", "unified", "common"])

    if is_portfolio:
        selected_projects = list(all_projects.values())
        print("\n" + "=" * 70)
        print(f"RUNNING UNIFIED MULTI-PROJECT DASHBOARD FOR ALL PROJECTS ({len(selected_projects)} CORPORATIONS)")
        print("=" * 70)
    else:
        if target_key not in all_projects:
            print(f"Error: Unknown project '{target_key}'. Configured projects: {', '.join(all_projects.keys())}")
            return False
        selected_projects = [all_projects[target_key]]
        print("\n" + "=" * 70)
        print(f"RUNNING REAL-TIME DASHBOARD FOR {selected_projects[0]['display_name'].upper()}")
        print("=" * 70)

    # Ingest data for all selected projects
    data_list = []
    for p_cfg in selected_projects:
        d = process_project_telemetry(p_cfg, live=live, portal_base=portal_base)
        data_list.append(d)

    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")

    # Destination paths
    if is_portfolio:
        output_folder = BASE_DIR
        excel_filename = f"schnell_smart_lights_portfolio_report_{date_str}.xlsx"
        report_title = f"Multi-Project Operations Portfolio ({len(data_list)} Corporations)"
    else:
        d = data_list[0]
        output_folder = d["project_config"]["folder"]
        excel_filename = f"{d['project_config']['name'].lower()}_nc_aging_report_{date_str}.xlsx"
        report_title = d["project_config"]["display_name"]

    excel_path = os.path.join(output_folder, excel_filename)
    html_path = os.path.join(output_folder, "mail_dashboard.html")

    # Generate Excel
    create_excel_report(data_list, excel_path, is_portfolio=is_portfolio)

    # Generate HTML
    html_content = generate_html_email(data_list, excel_filename, is_portfolio=is_portfolio)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Generated Real-Time HTML Dashboard: {html_path}")

    # Also keep a copy inside each project folder if portfolio
    if is_portfolio:
        for d in data_list:
            p_html = os.path.join(d["project_config"]["folder"], "mail_dashboard.html")
            with open(p_html, "w", encoding="utf-8") as f:
                f.write(html_content)

    # Console Summary
    tot_p = sum(d["total_panels"] for d in data_list)
    tot_on = sum(d["on_count"] for d in data_list)
    tot_off = sum(d["off_count"] for d in data_list)
    tot_nc = sum(d["nc_count"] for d in data_list)
    hlth = round(((tot_p - tot_nc) / tot_p * 100), 2) if tot_p > 0 else 0.0

    print("\n" + "-" * 55)
    print(f"EXECUTION SUMMARY ({report_title.upper()})")
    print("-" * 55)
    print(f"Timestamp          : {data_list[0]['timestamp']}")
    print(f"Total Fleet Panels : {tot_p:,}")
    print(f"Relay ON           : {tot_on:,}")
    print(f"Relay OFF          : {tot_off:,}")
    print(f"No Comm (NC)       : {tot_nc:,} ({round(tot_nc/tot_p*100, 1)}% Deficit)")
    print(f"Comm Health Rate   : {hlth}%")

    # Dispatch email if requested
    if send_email:
        p_key_tag = None if is_portfolio else selected_projects[0]["key"]
        send_dashboard_email(html_content, excel_path, report_title, custom_recipients, project_key=p_key_tag)
    else:
        print("\n[Ready to Send] Re-run with '--send' flag to dispatch email with Excel attachment via SMTP.")

    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Schnell Smart Lights Universal Dashboard Generator for All Projects")
    parser.add_argument("--project", default="all", help="Project key ('all', 'trichy', 'vellore', etc.)")
    parser.add_argument("--send", action="store_true", help="Send email report via SMTP")
    parser.add_argument("--to", default=None, help="Comma-separated recipient email addresses")
    parser.add_argument("--cached", action="store_true", help="Use saved CSV cache instead of live portal scraping")

    args = parser.parse_args()
    run_common_generator(project=args.project, send_email=args.send, custom_recipients=args.to, live=not args.cached)
