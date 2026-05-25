# Live Dashboard — One-time Setup

## What's built

| File | Purpose |
|------|---------|
| `index.html` | Dashboard — reads `data.json` dynamically |
| `data.json` | Auto-generated data snapshot |
| `update_data.py` | Fetches Excel from SharePoint → writes `data.json` |
| `.github/workflows/update.yml` | Runs `update_data.py` hourly (Mon–Fri 9am–7pm ET) |

**Live URL:** https://pedrotorresbasanta.github.io/peak-power-vt-dashboard/

---

## Step 1 — Azure App Registration (one-time, ~5 min)

1. Go to **portal.azure.com** → Azure Active Directory → App registrations → **New registration**
2. Name: `VT Dashboard` · Account type: *Accounts in this organizational directory only*
3. Click **Register**
4. Copy the **Application (client) ID** and **Directory (tenant) ID** — you'll need them below
5. Go to **Certificates & secrets** → New client secret → Add
6. Copy the **Value** immediately (it disappears after you leave)
7. Go to **API permissions** → Add a permission → Microsoft Graph → Application permissions
   - Add `Sites.Read.All` (or `Files.Read.All`)
8. Click **Grant admin consent**

---

## Step 2 — Add GitHub Secrets

Go to: https://github.com/pedrotorresbasanta/peak-power-vt-dashboard/settings/secrets/actions

Add three secrets:

| Secret name | Value |
|-------------|-------|
| `AZURE_TENANT_ID` | Directory (tenant) ID from Step 1 |
| `AZURE_CLIENT_ID` | Application (client) ID from Step 1 |
| `AZURE_CLIENT_SECRET` | Client secret value from Step 1 |

---

## Step 3 — Trigger first run

Go to: https://github.com/pedrotorresbasanta/peak-power-vt-dashboard/actions  
Select **Refresh Dashboard Data** → **Run workflow** → **Run workflow**

After ~30 seconds `data.json` will be updated with live SharePoint data and the dashboard refreshes.

---

## Updating the Excel

Just save the file normally in SharePoint — the Actions workflow picks it up automatically on its next hourly run.  
To force an immediate update: Actions → Refresh Dashboard Data → Run workflow.

---

## Running locally

```bash
LOCAL_XLSX_PATH=/path/to/VT_Weekly_Template_FINAL.xlsx python update_data.py
open index.html   # needs a local server for fetch() to work
python -m http.server 8000  # then open http://localhost:8000
```
