# Trichy Smart Lights Automation Workflow

This directory contains the automated data collection and reporting workflow for the **Trichy Smart Streetlights Project** on [TN Smart Lights](https://tnsmartlights.in).

## Configuration Details

- **Portal URL**: `https://tnsmartlights.in`
- **Username**: `tycadmin`
- **Password**: `tyc12249`
- **Project ID**: `7` (TRICHY)
- **Zones**: 4 Zones (Zone 1, Zone 2, Zone 3, Zone 4)
- **Wards**: 64 Wards
- **Total Devices**: 1,807 panels (API) / 1,809 panels (Liveview)

---

## Scripts & Operations

### 1. `login_test.py`
Connects to the portal, authenticates as `tycadmin`, traverses all 4 zones and 64 wards, loads map points, and dumps project structure:
```bash
python login_test.py
```
**Outputs**:
- `project_7_data.json`: Ward and Zone metadata mapping.
- `project_7_devices.json`: Full device inventory with coordinates, status, and IDs.
- `nc_devices.json`: Filtered list of Not-Communicating (status = 2) devices.

---

### 2. `relay_check.py`
Queries the live dashboard map API status and launches Playwright to scrape the complete real-time liveview table:
```bash
python relay_check.py
```
**Outputs**:
- `trichy_relay_status.csv`: Live snapshot of all panels (UID, Zone, Ward, Relay State, Dashboard Status, Last Update, Previous Update, Fault).
- Terminal summary of Relay ON, Relay OFF, Unknown, and NC percentages.

---

### 3. `get_nc_reports.py`
Cross-references live dashboard status with liveview timestamps to calculate offline durations:
```bash
python get_nc_reports.py
```
**Outputs**:
- `trichy_nc_over_7_days.csv`: Panels offline for longer than 7 days, sorted descending by offline duration.
- `trichy_all_nc_panels.csv`: All known non-communicating panels with last update timestamp and offline duration.

---

### 4. Mail Dashboard & Excel Report Generation
The dashboard and Excel report are generated using the common universal engine at root:
```bash
python ../generate_dashboard.py --project trichy
```
To send via SMTP:
```bash
python ../generate_dashboard.py --project trichy --send
```
**Outputs**:
- `mail_dashboard.html`: Email-ready rich dashboard (Total, Relay ON, Relay OFF, NC, Zone Breakdown, Aging Cards).
- `trichy_nc_aging_report_*.xlsx`: Formatted Excel report with Days in NC, Zone Summary, and Fleet Status.
