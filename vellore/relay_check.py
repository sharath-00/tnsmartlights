import csv
import time
import requests
from collections import Counter
from playwright.sync_api import sync_playwright

# ============================================================
# CONFIGURATION
# ============================================================

USERNAME = "vcadmin"
PASSWORD = "vel5555"
OUTPUT_CSV = "vellore_relay_status.csv"

# ============================================================
# PLAYWRIGHT AUTOMATION
# ============================================================

def run_relay_check():
    all_records = []
    dashboard_status_map = {}
    
    print("Fetching dashboard status map via API...")
    try:
        session = requests.Session()
        login_data = {"Username": USERNAME, "password": PASSWORD, "send": "Sign in"}
        session.post("https://tnsmartlights.in/slm/home/", data=login_data, allow_redirects=True)
        zones = session.post("https://tnsmartlights.in/slm/home/loadzone/", data="3").json()
        for zone in zones:
            wards = session.post("https://tnsmartlights.in/slm/home/loadward/", data=f"3,{zone[0]}").json()
            for ward in wards:
                try:
                    devices = session.post("https://tnsmartlights.in/slm/home/loadmap/", data=f"3!zone!{zone[0]}!{ward[0]}").json()
                    for d in devices:
                        if "uid" in d and "status" in d:
                            dashboard_status_map[d["uid"]] = str(d["status"])
                except Exception:
                    pass
        print(f"Successfully loaded dashboard status for {len(dashboard_status_map)} devices.")
    except Exception as e:
        print(f"Failed to fetch dashboard map data: {e}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("\nLogging in to live view...")
        page.goto("https://tnsmartlights.in/slm/")
        page.fill('input[name="Username"]', USERNAME)
        page.fill('input[name="password"]', PASSWORD)
        page.click('input[name="send"]')
        
        print("Waiting for dashboard to load...")
        page.wait_for_url("https://tnsmartlights.in/slm/home/")
        
        print("Loading liveview map (fetching all 1000+ panels instantly)...")
        page.goto("https://tnsmartlights.in/slm/home/liveview/")
        
        print("Waiting for data table to populate from server (this takes ~10 seconds)...")
        page.wait_for_selector("#rlclist tbody tr", timeout=30000)
        
        rows = page.locator("#rlclist tbody tr")
        total_rows = rows.count()
        print(f"Successfully loaded {total_rows} devices!")
        
        print("Parsing live data...")
        for i in range(total_rows):
            row = rows.nth(i)
            cells = row.locator("td")
            
            try:
                # Based on the HTML DOM mapping
                uid = cells.nth(3).inner_text().strip()
                zone = cells.nth(4).inner_text().strip()
                ward = cells.nth(5).inner_text().strip()
                
                # Column 19 is Last Update, 20 is Previous Update
                last_update = cells.nth(19).inner_text().strip().split('\n')[0]
                previous_update = cells.nth(20).inner_text().strip().split('\n')[0]
                
                # Column 22 is Relay Status ("ON", "OFF", "NO COMM")
                state = cells.nth(22).inner_text().strip()
                
                # Column 23 is Fault (e.g. TAMPER, Power-Fail)
                fault = cells.nth(23).inner_text().strip().split('\n')[0]
                
                if not state:
                    state = "UNKNOWN"

                dashboard_status = dashboard_status_map.get(uid, "UNKNOWN")

                record = {
                    "uid": uid,
                    "zone": zone,
                    "ward": ward,
                    "state": state,
                    "dashboard_status": dashboard_status,
                    "last_update": last_update,
                    "previous_update": previous_update,
                    "fault": fault,
                }
                
                all_records.append(record)
                
            except Exception as e:
                pass
                
        browser.close()

    # ============================================================
    # SAVE CSV
    # ============================================================
    
    if not all_records:
        print("\nNo records to save.")
        return
        
    print(f"\nSaving {len(all_records)} records to {OUTPUT_CSV}...")
    
    fieldnames = [
        "uid",
        "zone",
        "ward",
        "state",
        "dashboard_status",
        "last_update",
        "previous_update",
        "fault"
    ]
    
    with open(OUTPUT_CSV, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in all_records:
            writer.writerow(record)
            
    # ============================================================
    # COUNT RELAY STATES (DASHBOARD LOGIC)
    # ============================================================

    # Dashboard logic: 
    # Status 1 = ON
    # Status 0 = OFF
    # Status 2 = NO COMM
    on_count = sum(1 for r in all_records if r["dashboard_status"] == "1")
    off_count = sum(1 for r in all_records if r["dashboard_status"] == "0")
    no_comm_count = sum(1 for r in all_records if r["dashboard_status"] == "2")
    unknown_count = len(all_records) - (on_count + off_count + no_comm_count)

    # ============================================================
    # SUMMARY
    # ============================================================

    print("\n")
    print("========================================")
    print("       VELLORE RELAY SUMMARY")
    print("========================================")
    print(f"Total live responses : {len(all_records)}")
    print(f"Relay ON             : {on_count}")
    print(f"Relay OFF            : {off_count}")
    print(f"Unknown              : {unknown_count}")
    print(f"No Comm (NC)         : {no_comm_count}")
    print("========================================")

    # ============================================================
    # PERCENTAGES
    # ============================================================

    total_known = on_count + off_count
    if total_known > 0:
        on_percentage = (on_count / total_known) * 100
        off_percentage = (off_count / total_known) * 100
        
        print(f"ON percentage  : {on_percentage:.2f}%")
        print(f"OFF percentage : {off_percentage:.2f}%")

    print("\nCompleted.")

if __name__ == "__main__":
    run_relay_check()