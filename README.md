# Incentive Dashboard

Internal Google Apps Script web app that gives trainers a personalized view of their weekly task completions, incentive pay tiers, consistency multipliers, and a ranked leaderboard. Data is read from a Google Sheet; optional scripts sync Airtable bases into other sheets for production and roster data.

---

## Repository structure

| File | Purpose |
|------|--------|
| **Code.js** | Backend for the dashboard: web app entry (`doGet`), main data endpoint (`getDashboardData`), incentive tier logic, leaderboard building, and week/date helpers. Reads from the `Gold_Testing` sheet and enforces access by email allow-list. |
| **Index.html** | Single-page web UI: “Momentum Dashboard” and “Progress Leaderboard” layout, styles, and client-side logic that calls `google.script.run.getDashboardData()` and renders the response (completions, pay progress, multiplier badge, earnings, leaderboard). |
| **Hubstaff.js** | Syncs an Airtable base into a Google Sheet named **Hubstaff** (incremental by “Created time”). Use for production/task data. Credentials should be stored in [PropertiesService](https://developers.google.com/apps-script/reference/properties) (Script Properties). |
| **Roster.js** | Syncs the Airtable **Agent Roster** table into a Google Sheet named **Roster** (full overwrite). Use for trainer/agent roster. Credentials should be stored in PropertiesService. |
| **appsscript.json** | Apps Script project config: timezone, web app access (`DOMAIN`), execute-as (`USER_ACCESSING`), exception logging (Stackdriver), and V8 runtime. |
| **.clasp.json** | [clasp](https://github.com/google/clasp) config for local development and deployment (script ID, file extensions). Do not commit credentials; keep `creds.json` out of the repo (see `.gitignore`). |
| **.gitignore** | Ignores `.clasp.json`, `node_modules/`, `.DS_Store`, and `creds.json` so secrets and local artifacts stay out of version control. |
| **Security_Review.md** | Security review of the Incentive Dashboard: feature overview, threat model, trust boundaries, OWASP/MITRE-aligned checklist, and consolidated risk list. |

---

## Documentation and resource

- **Notion:** [Incentive Dashboard – Notion](https://www.notion.so/invisibletech/OpenAI-Incentive-Dashboard-Security-Review-2fb82d3947a6805b88d0c30524a9e42b?source=copy_link) 

