import requests
from collections import Counter

BASE_URL = "https://tnsmartlights.in"
PROJECT_ID = "3"

USERNAME = "vcadmin"
PASSWORD = "vel5555"

session = requests.Session()

headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": BASE_URL + "/slm/"
}


# --------------------------------------------------
# 1. LOGIN
# --------------------------------------------------

print("Logging in...")

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
    exit()

print("Login successful.")


# --------------------------------------------------
# 2. GET ALL ZONES
# --------------------------------------------------

print("\nFetching zones...")

response = session.post(
    BASE_URL + "/slm/home/loadzone/",
    data=PROJECT_ID,
    headers=headers
)

zones = response.json()

print("Total zones:", len(zones))

for zone in zones:
    print("Zone ID:", zone[0], "Zone Name:", zone[1])


# --------------------------------------------------
# 3. FETCH ALL WARDS + MAP DATA
# --------------------------------------------------

status_counter = Counter()

total_devices = 0
total_wards = 0

all_devices = []


for zone in zones:

    zone_id = zone[0]
    zone_name = zone[1]

    print(f"\nProcessing Zone {zone_name}...")

    # Get wards for this zone
    response = session.post(
        BASE_URL + "/slm/home/loadward/",
        data=f"{PROJECT_ID},{zone_id}",
        headers=headers
    )

    wards = response.json()

    print("Wards:", len(wards))

    for ward in wards:

        ward_id = ward[0]
        ward_name = ward[1]

        total_wards += 1

        print(
            f"  Processing Ward {ward_name} "
            f"(ID: {ward_id})"
        )

        # ------------------------------------------
        # LOAD MAP DATA
        # ------------------------------------------

        payload = f"{PROJECT_ID}!zone!{zone_id}!{ward_id}"

        response = session.post(
            BASE_URL + "/slm/home/loadmap/",
            data=payload,
            headers=headers
        )

        # Make sure response is JSON
        try:
            devices = response.json()
        except Exception:
            print("    ERROR: API did not return JSON")
            print("    Status:", response.status_code)
            continue

        print("    Devices:", len(devices))

        # ------------------------------------------
        # COUNT STATUS
        # ------------------------------------------

        for device in devices:

            status = str(device.get("status", ""))

            status_counter[status] += 1
            total_devices += 1

            all_devices.append({
                "zone_id": zone_id,
                "zone_name": zone_name,
                "ward_id": ward_id,
                "ward_name": ward_name,
                "device_id": device.get("id"),
                "uid": device.get("uid"),
                "latitude": device.get("lat"),
                "longitude": device.get("lng"),
                "status": status
            })


# --------------------------------------------------
# 4. DISPLAY RESULTS
# --------------------------------------------------

print("\n" + "=" * 50)
print("FINAL RESULT")
print("=" * 50)

print("Project ID       :", PROJECT_ID)
print("Total Zones      :", len(zones))
print("Total Wards      :", total_wards)
print("Total Devices    :", total_devices)

print("\nSTATUS COUNTS")
print("-" * 30)

for status, count in status_counter.items():
    print(f"Status {status}: {count}")


# --------------------------------------------------
# 5. DISPLAY POSSIBLE NOT COMMUNICATION COUNT
# --------------------------------------------------

print("\n" + "=" * 50)
print("NOT COMMUNICATION")
print("=" * 50)

# Based on the data we have seen so far,
# status 2 appears in the loadmap response.
# DO NOT assume this is definitely "Not Communication"
# until confirmed from the website's status mapping.

not_communication = status_counter.get("2", 0)

print("Status 2 count:", not_communication)