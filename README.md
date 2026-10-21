# Schnell Energy Equipments • Universal Smart Lights Fleet Operations

Universal automation workflows, real-time relay telemetry checks, non-communicating (NC) panel aging analytics, executive mail dashboards, and formatted Excel reporting for Tamil Nadu Smart Lights deployments ([tnsmartlights.in](https://tnsmartlights.in)).

---

## Workspace Structure

```
d:\Schnell\tnsmartlights\
├── .github/workflows/
│   ├── trichy_smartlights_dispatch.yml   # GitHub Actions automated workflow for Trichy
│   └── vellore_smartlights_dispatch.yml  # GitHub Actions automated workflow for Vellore
├── generate_dashboard.py                 # Universal Common Dashboard & Telemetry Generator
├── mail_dashboard_generator.py           # Compatibility wrapper
├── requirements.txt                      # Project dependencies
├── GITHUB_ACTIONS_SETUP.md               # GitHub Actions setup and secrets guide
├── .env                                  # Common credentials & SMTP configuration
├── .env.example                          # Environment template
├── trichy\                               # Trichy Smart Lights (Project 7)
│   ├── login_test.py                     # Zone, ward & device discovery
│   ├── relay_check.py                    # Real-time relay & liveview check
│   ├── get_nc_reports.py                 # Offline duration & NC reports
│   ├── mail_dashboard.html               # Generated email HTML dashboard
│   ├── trichy_nc_aging_report_*.xlsx     # Formatted Excel report (Days in NC)
│   ├── trichy_relay_status.csv           # Liveview data export
│   └── README.md                         # Trichy documentation
└── vellore\                              # Vellore Smart Lights (Project 3)
    ├── login_test.py                     # Zone, ward & device discovery
    ├── relay_check.py                    # Real-time relay & liveview check
    ├── get_nc_reports.py                 # Offline duration & NC reports
    ├── mail_dashboard.html               # Generated email HTML dashboard
    ├── vellore_nc_aging_report_*.xlsx    # Formatted Excel report (Days in NC)
    ├── vellore_relay_status.csv          # Liveview data export
    └── README.md                         # Vellore documentation
```

---

## GitHub Actions Workflows

Two automated CI/CD workflows are configured under `.github/workflows/`:
1. **[trichy_smartlights_dispatch.yml](file:///.github/workflows/trichy_smartlights_dispatch.yml)**: Automated Trichy telemetry extraction, dashboard generation, and email dispatch.
2. **[vellore_smartlights_dispatch.yml](file:///.github/workflows/vellore_smartlights_dispatch.yml)**: Automated Vellore telemetry extraction, dashboard generation, and email dispatch.

For complete setup instructions (secrets, schedules, and manual triggers), see **[GITHUB_ACTIONS_SETUP.md](file:///d:/Schnell/tnsmartlights/GITHUB_ACTIONS_SETUP.md)**.

---

## Universal Common Dashboard Generator (`generate_dashboard.py`)

A single, dynamic generator that manages **all projects**:

### 1. Unified Multi-Project Portfolio Dashboard (Default)
Generates an executive portfolio dashboard combining **all active projects** (Trichy + Vellore) with aggregated KPI cards, project-wise comparison table, individual zone tables, and a unified multi-sheet Excel workbook:
```bash
python generate_dashboard.py
# or explicitly:
python generate_dashboard.py --project all
```

### 2. Dedicated Single-Project Dashboard
Generates a dedicated dashboard and Excel workbook for a specific project:
```bash
# Trichy Corporation
python generate_dashboard.py --project trichy

# Vellore Corporation
python generate_dashboard.py --project vellore
```

### 3. Send Email via SMTP (With Attached Excel Workbook)
Add `--send` to dispatch the generated HTML dashboard with the Excel report attached:
```bash
# Send portfolio dashboard (all projects)
python generate_dashboard.py --send

# Send Trichy dashboard
python generate_dashboard.py --project trichy --send

# Send Vellore dashboard
python generate_dashboard.py --project vellore --send

# Send to custom recipient on the fly
python generate_dashboard.py --project trichy --send --to "recipient@schnellenergy.com"
```

### 4. Fast Cached Mode (Offline Re-render)
If you want to re-render using existing data without hitting the portal, pass `--cached`:
```bash
python generate_dashboard.py --project all --cached
```

---

## Adding New Projects in the Future

To add any new city/corporation (e.g. Madurai, Salem, Coimbatore), simply add its credentials to `.env`:
```env
MADURAI_NAME=Madurai Municipal Corporation
MADURAI_PROJECT_ID=12
MADURAI_USERNAME=mdadmin
MADURAI_PASSWORD=your_password
```
`generate_dashboard.py` will **automatically discover** it and include it in both single and multi-project reporting without changing any Python code.
