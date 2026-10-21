import csv
from playwright.sync_api import sync_playwright
from datetime import datetime
import json
import os
import requests

USERNAME = "tycadmin"
PASSWORD = "tyc12249"
PROJECT_ID = "7"
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_nc_reports():
    print("Fetching live dashboard status via API to find NC panels in Trichy (Project 7)...")
    nc_uids = set()
    try:
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://tnsmartlights.in/slm/"
        }
        login_data = {"Username": USERNAME, "password": PASSWORD, "send": "Sign in"}
        session.post("https://tnsmartlights.in/slm/home/", data=login_data, headers=headers, allow_redirects=True)
        zones = session.post("https://tnsmartlights.in/slm/home/loadzone/", data=PROJECT_ID, headers=headers).json()
        
        for zone in zones:
            wards = session.post("https://tnsmartlights.in/slm/home/loadward/", data=f"{PROJECT_ID},{zone[0]}", headers=headers).json()
            for ward in wards:
                try:
                    devices = session.post("https://tnsmartlights.in/slm/home/loadmap/", data=f"{PROJECT_ID}!zone!{zone[0]}!{ward[0]}", headers=headers).json()
                    for d in devices:
                        if "uid" in d and "status" in d:
                            if str(d["status"]) == "2":
                                nc_uids.add(d["uid"])
                except Exception:
                    pass
        print(f"Successfully loaded live map data. Found {len(nc_uids)} live NC devices.")
    except Exception as e:
        print(f"Failed to fetch live dashboard map data: {e}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("Navigating to tnsmartlights.in...")
        page.goto("https://tnsmartlights.in/slm/")
        
        print("Logging in...")
        page.fill('input[name="Username"]', USERNAME)
        page.fill('input[name="password"]', PASSWORD)
        page.click('input[name="send"]')
        
        print("Waiting for dashboard to load...")
        page.wait_for_url("https://tnsmartlights.in/slm/home/")
        
        print("Going to liveview...")
        page.goto("https://tnsmartlights.in/slm/home/liveview/")
        
        print("Waiting for data table to populate (this might take 10-15 seconds)...")
        page.wait_for_selector("#rlclist tbody tr", timeout=30000)
        
        rows = page.locator("#rlclist tbody tr")
        total_rows = rows.count()
        print(f"Successfully loaded {total_rows} devices from the liveview server!")
        
        all_nc_panels = []
        nc_over_7_panels = []
        current_time = datetime.now()

        print("\nAnalyzing offline time for all panels...")
        for i in range(total_rows):
            row = rows.nth(i)
            cells = row.locator("td")
            
            try:
                uid = cells.nth(3).inner_text().strip()
                zone = cells.nth(4).inner_text().strip()
                ward = cells.nth(5).inner_text().strip()
                location = cells.nth(6).inner_text().strip()
                last_update_str = cells.nth(19).inner_text().strip().split('\n')[0]
                state = cells.nth(22).inner_text().strip()
                fault = cells.nth(23).inner_text().strip().split('\n')[0]
                
                # Calculate date diff if valid
                date_diff_days = -1
                if last_update_str and last_update_str.lower() != "null" and last_update_str != "-":
                    dt = datetime.strptime(last_update_str.strip(), "%Y-%m-%d %H:%M:%S")
                    date_diff_days = (current_time - dt).days + (current_time - dt).seconds / 86400.0
                
                panel_data = {
                    "UID": uid,
                    "Zone": zone,
                    "Ward": ward,
                    "Location": location,
                    "State": state,
                    "Fault": fault,
                    "Last Update": last_update_str if date_diff_days != -1 else "No Data",
                    "Days Offline": round(date_diff_days, 2) if date_diff_days != -1 else 0
                }

                # If UID is in known NC panels list
                if uid in nc_uids:
                    all_nc_panels.append(panel_data)
                
                # If panel has been offline for > 7 days (across all panels)
                if date_diff_days > 7.0:
                    nc_over_7_panels.append(panel_data)

            except Exception:
                pass
        
        browser.close()
        
        # Helper to format time
        def format_time(days):
            if days == 0: return "Unknown"
            if days < 1:
                hours = round(days * 24, 1)
                if hours < 1:
                    mins = round(days * 24 * 60)
                    return f"{mins} mins"
                return f"{hours} hours"
            return f"{days} days"

        # Sort both lists descending by offline time
        all_nc_panels.sort(key=lambda x: x["Days Offline"], reverse=True)
        nc_over_7_panels.sort(key=lambda x: x["Days Offline"], reverse=True)
        
        print(f"\n========================================")
        print(f"REPORT 1: PANELS OFFLINE FOR > 7 DAYS (Total: {len(nc_over_7_panels)})")
        print(f"========================================")
        for panel in nc_over_7_panels[:50]:  # Show top 50 in terminal
            time_str = format_time(panel['Days Offline'])
            print(f"UID: {panel['UID']:<15} | Ward: {panel['Ward']:<5} | Last Update: {panel['Last Update']:<20} | Offline for: {time_str}")
        if len(nc_over_7_panels) > 50:
            print(f"... and {len(nc_over_7_panels) - 50} more panels (full list saved to CSV).")

        print(f"\n========================================")
        print(f"REPORT 2: ALL KNOWN NC PANELS (Total: {len(all_nc_panels)})")
        print(f"========================================")
        for panel in all_nc_panels[:50]:  # Show top 50 in terminal
            time_str = format_time(panel['Days Offline'])
            print(f"UID: {panel['UID']:<15} | Ward: {panel['Ward']:<5} | Last Update: {panel['Last Update']:<20} | Offline for: {time_str}")
        if len(all_nc_panels) > 50:
            print(f"... and {len(all_nc_panels) - 50} more panels (full list saved to CSV).")

        # Save CSV reports
        csv_over_7 = os.path.join(CURRENT_DIR, "trichy_nc_over_7_days.csv")
        csv_all_nc = os.path.join(CURRENT_DIR, "trichy_all_nc_panels.csv")
        
        fieldnames = ["UID", "Zone", "Ward", "Location", "State", "Fault", "Last Update", "Days Offline"]
        
        with open(csv_over_7, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for p in nc_over_7_panels:
                writer.writerow(p)
                
        with open(csv_all_nc, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for p in all_nc_panels:
                writer.writerow(p)

        print(f"\nSaved Report 1 to: {csv_over_7}")
        print(f"Saved Report 2 to: {csv_all_nc}")

if __name__ == "__main__":
    print("Starting Trichy NC reports analysis...")
    get_nc_reports()
