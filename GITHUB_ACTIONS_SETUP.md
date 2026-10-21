# GitHub Actions Workflows Setup Guide

Automated daily execution and email dispatch for **Trichy** and **Vellore** smart lighting projects using GitHub Actions.

---

## 1. Workflows Included

| Workflow | File | Triggers | Default Schedule (IST) |
| :--- | :--- | :--- | :--- |
| **Trichy Smart Lights** | [trichy_smartlights_dispatch.yml](file:///.github/workflows/trichy_smartlights_dispatch.yml) | Manual, Schedule, Webhook | 06:00 AM IST & 07:00 PM IST |
| **Vellore Smart Lights** | [vellore_smartlights_dispatch.yml](file:///.github/workflows/vellore_smartlights_dispatch.yml) | Manual, Schedule, Webhook | 06:15 AM IST & 07:15 PM IST |

---

## 2. Setting Up GitHub Repository Secrets

In your GitHub repository, go to **Settings** > **Secrets and variables** > **Actions** > **New repository secret**:

### Required Email (SMTP) Secrets:
- `SMTP_USERNAME`: `vinoth.joel@schnellenergy.com` (or your company email)
- `SMTP_PASSWORD`: Your Gmail App Password (16 characters, e.g. `ufedzyzpdtlzyjlx`)
- `SENDER_EMAIL`: `vinoth.joel@schnellenergy.com`
- `RECIPIENTS`: Comma-separated email list (e.g. `management@schnellenergy.com,tl@schnellenergy.com`)

### Optional Portal Secrets (Defaults are already baked in):
- `TRICHY_USERNAME`: `tycadmin`
- `TRICHY_PASSWORD`: `tyc12249`
- `VELLORE_USERNAME`: `vcadmin`
- `VELLORE_PASSWORD`: `vel5555`

---

## 3. How to Trigger Workflows

### Option A: Manual Trigger via GitHub UI
1. Navigate to the **Actions** tab in your repository.
2. Under "Workflows", select either **Trichy Smart Lights Automated Dashboard Dispatch** or **Vellore Smart Lights Automated Dashboard Dispatch**.
3. Click the **Run workflow** dropdown on the right:
   - Check/uncheck `send_email` (default: true).
   - Optionally enter a specific `recipient` override.
4. Click **Run workflow**.

### Option B: Automated Daily Schedules
- **Morning Run (06:00 AM IST)**: Verifies daytime streetlight shutoff (Relay OFF validation) and identifies new NC panels.
- **Evening Run (07:00 PM IST)**: Verifies nighttime streetlight turn-on (Relay ON validation) and power feeder health.

### Option C: Webhook / External API Trigger (cURL / Postman / AppSheet)

#### 1. Trigger BOTH Trichy & Vellore Emails Simultaneously (Single Request)
```bash
curl -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: Bearer YOUR_GITHUB_PERSONAL_ACCESS_TOKEN" \
  https://api.github.com/repos/sharath-00/tnsmartlights/dispatches \
  -d '{"event_type": "all_dashboards_dispatch"}'
```

#### 2. Trigger Trichy Only
```bash
curl -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: Bearer YOUR_GITHUB_PERSONAL_ACCESS_TOKEN" \
  https://api.github.com/repos/sharath-00/tnsmartlights/dispatches \
  -d '{"event_type": "trichy_dashboard_dispatch"}'
```

#### 3. Trigger Vellore Only
```bash
curl -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: Bearer YOUR_GITHUB_PERSONAL_ACCESS_TOKEN" \
  https://api.github.com/repos/sharath-00/tnsmartlights/dispatches \
  -d '{"event_type": "vellore_dashboard_dispatch"}'
```

#### 4. Optional: Override Recipient via `client_payload`
```bash
curl -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: Bearer YOUR_GITHUB_PERSONAL_ACCESS_TOKEN" \
  https://api.github.com/repos/sharath-00/tnsmartlights/dispatches \
  -d '{"event_type": "all_dashboards_dispatch", "client_payload": {"recipient": "your_email@example.com"}}'
```

---

## 4. Run Artifacts

Every workflow run automatically saves:
- `mail_dashboard.html`
- Timestamped `.xlsx` report (with Days in NC, Zone breakdown, and Fleet status)
- Generated `.csv` telemetry files

These can be downloaded directly from the **Actions run summary page** for 30 days.
