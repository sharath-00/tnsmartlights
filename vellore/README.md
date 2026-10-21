# Vellore Smart Lights Automation Workflow

This directory contains the automated data collection and reporting workflow for the **Vellore Smart Streetlights Project** on [TN Smart Lights](https://tnsmartlights.in).

## Configuration Details

- **Portal URL**: `https://tnsmartlights.in`
- **Username**: `vcadmin`
- **Password**: `vel5555`
- **Project ID**: `3` (VELLORE)

---

## Scripts & Operations

### 1. `login_test.py`
Connects to the portal, authenticates as `vcadmin`, traverses zones and wards, and displays device status counts:
```bash
python login_test.py
```
**Data files**:
- `project_3_data.json`: Zone and ward mapping.
- `project_3_devices.json`: Full device details.
- `nc_devices.json`: Filtered list of Not-Communicating (NC) devices.

---

### 2. `relay_check.py`
Queries dashboard status map and launches Playwright to fetch liveview table:
```bash
python relay_check.py
```
**Output**:
- `vellore_relay_status.csv`: Live device relay status and faults.

---

### 3. `get_nc_reports.py`
Queries live map status and liveview data to identify panels offline > 7 days and all NC panels:
```bash
python get_nc_reports.py
```

---

### 4. Mail Dashboard & Excel Report Generation
The dashboard and Excel report are generated using the common universal engine at root:
```bash
python ../generate_dashboard.py --project vellore
```
To send via SMTP:
```bash
python ../generate_dashboard.py --project vellore --send
```
**Outputs**:
- `mail_dashboard.html`: Email-ready rich dashboard (Total, Relay ON, Relay OFF, NC, Zone Breakdown, Aging Cards).
- `vellore_nc_aging_report_*.xlsx`: Formatted Excel report with Days in NC, Zone Summary, and Fleet Status.
