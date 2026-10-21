import json
import os
import requests
from collections import Counter

BASE_URL = "https://tnsmartlights.in"
PROJECT_ID = "7"

USERNAME = "tycadmin"
PASSWORD = "tyc12249"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

session = requests.Session()

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": BASE_URL + "/slm/"
}


# --------------------------------------------------
# 1. LOGIN
# --------------------------------------------------

print("Logging in to TN Smart Lights (Trichy)...")

session.get(
    BASE_URL + "/slm/",
    headers=headers
)

login_data = {
    "Username": USERNAME,
    "password": PASSWORD,
    "send": "Sign in"
}

login_response = session.post(
    BASE_URL + "/slm/home/",
    data=login_data,
    headers=headers,
    allow_redirects=True
)

if login_response.url != BASE_URL + "/slm/home/":
    print("Login may have failed.")
    print("Final URL:", login_response.url)
    exit(1)

print("Login successful.")


# --------------------------------------------------
# 2. GET ALL ZONES
# --------------------------------------------------

print("\nFetching zones for Project ID:", PROJECT_ID)

response = session.post(
    BASE_URL + "/slm/home/loadzone/",
    data=PROJECT_ID,
    headers=headers
)

zones = response.json()

print("Total zones:", len(zones))

for zone in zones:
    print(f"Zone ID: {zone[0]}, Zone Name: {zone[1]}")


# --------------------------------------------------
# 3. FETCH ALL WARDS + MAP DATA
# --------------------------------------------------

status_counter = Counter()

total_devices = 0
total_wards = 0

all_wards_data = []
all_devices = []
nc_devices = []


for zone in zones:
    zone_id = zone[0]
    zone_name = zone[1]

    print(f"\nProcessing Zone {zone_name} (ID: {zone_id})...")

    # Get wards for this zone
    response = session.post(
        BASE_URL + "/slm/home/loadward/",
        data=f"{PROJECT_ID},{zone_id}",
        headers=headers
    )

    wards = response.json()
    print("  Wards count:", len(wards))

    for ward in wards:
        ward_id = ward[0]
        ward_name = ward[1]

        total_wards += 1
        all_wards_data.append({
            "project_id": int(PROJECT_ID),
            "zone_id": zone_id,
            "zone_name": str(zone_name),
            "ward_id": ward_id,
            "ward_name": str(ward_name)
        })

        print(f"    Processing Ward {ward_name} (ID: {ward_id})...")

        # ------------------------------------------
        # LOAD MAP DATA
        # ------------------------------------------
        payload = f"{PROJECT_ID}!zone!{zone_id}!{ward_id}"

        response = session.post(
            BASE_URL + "/slm/home/loadmap/",
            data=payload,
            headers=headers
        )

        try:
            devices = response.json()
        except Exception:
            print("      ERROR: API did not return JSON")
            print("      Status:", response.status_code)
            continue

        print(f"      Devices: {len(devices)}")

        # ------------------------------------------
        # COUNT STATUS & STORE
        # ------------------------------------------
        for device in devices:
            status = str(device.get("status", ""))
            status_counter[status] += 1
            total_devices += 1

            uid = device.get("uid")
            dev_id = device.get("id")

            dev_record = {
                "project_id": int(PROJECT_ID),
                "zone_id": zone_id,
                "zone_name": str(zone_name),
                "ward_id": ward_id,
                "ward_name": str(ward_name),
                "device_id": str(dev_id),
                "uid": uid,
                "latitude": device.get("lat"),
                "longitude": device.get("lng"),
                "status": status
            }
            all_devices.append(dev_record)

            if status == "2":
                nc_devices.append({
                    "zone": str(zone_name),
                    "ward": str(ward_name),
                    "uid": uid,
                    "id": str(dev_id)
                })


# --------------------------------------------------
# 4. SAVE DATA ARTIFACTS
# --------------------------------------------------

wards_file = os.path.join(CURRENT_DIR, f"project_{PROJECT_ID}_data.json")
with open(wards_file, "w", encoding="utf-8") as f:
    json.dump(all_wards_data, f, indent=4)
print(f"\nSaved {len(all_wards_data)} wards to {wards_file}")

devices_file = os.path.join(CURRENT_DIR, f"project_{PROJECT_ID}_devices.json")
with open(devices_file, "w", encoding="utf-8") as f:
    json.dump(all_devices, f, indent=4)
print(f"Saved {len(all_devices)} devices to {devices_file}")

nc_file = os.path.join(CURRENT_DIR, "nc_devices.json")
with open(nc_file, "w", encoding="utf-8") as f:
    json.dump(nc_devices, f, indent=2)
print(f"Saved {len(nc_devices)} NC devices to {nc_file}")


# --------------------------------------------------
# 5. DISPLAY RESULTS
# --------------------------------------------------

print("\n" + "=" * 50)
print("FINAL RESULT (TRICHY)")
print("=" * 50)

print("Project ID       :", PROJECT_ID)
print("Total Zones      :", len(zones))
print("Total Wards      :", total_wards)
print("Total Devices    :", total_devices)

print("\nSTATUS COUNTS")
print("-" * 30)

for status, count in sorted(status_counter.items()):
    status_label = "OFF" if status == "0" else "ON" if status == "1" else "NC (No Comm)" if status == "2" else "Other"
    print(f"Status {status} ({status_label}): {count}")


# --------------------------------------------------
# 6. DISPLAY POSSIBLE NOT COMMUNICATION COUNT
# --------------------------------------------------

print("\n" + "=" * 50)
print("NOT COMMUNICATION")
print("=" * 50)

not_communication = status_counter.get("2", 0)
print("Status 2 (No Communication) count:", not_communication)
