# Incentive Dashboard

Internal Google Apps Script web app that gives trainers a personalized view of their weekly task completions, incentive pay tiers, consistency multipliers, and a ranked leaderboard. Data is read from a single Google Sheet (`Gold_Testing`).

---

## Repository structure

| File | Purpose |
|------|--------|
| **Code.js** | Backend for the dashboard: web app entry (`doGet`), main data endpoint (`getDashboardData`), incentive tier logic, leaderboard building, and week/date helpers. Reads from the `Gold_Testing` sheet and enforces access by email allow-list. |
| **Index.html** | Single-page web UI: “Momentum Dashboard” and “Progress Leaderboard” layout, styles, and client-side logic that calls `google.script.run.getDashboardData()` and renders the response (completions, pay progress, multiplier badge, earnings, leaderboard). |
| **appsscript.json** | Apps Script project config: timezone, web app access (`DOMAIN`), execute-as (`USER_ACCESSING`), exception logging (Stackdriver), and V8 runtime. |
| **.clasp.json** | [clasp](https://github.com/google/clasp) config for local development and deployment (script ID, file extensions). Do not commit credentials; keep `creds.json` out of the repo (see `.gitignore`). |
| **.gitignore** | Ignores `.clasp.json`, `node_modules/`, `.DS_Store`, and `creds.json` so secrets and local artifacts stay out of version control. |
| **Security_Review.md** | Security review of the Incentive Dashboard: feature overview, threat model, trust boundaries, OWASP/MITRE-aligned checklist, and consolidated risk list. |

---

## Where to find things

- **Run the web app:** Deploy as a web app from the [Apps Script editor](https://script.google.com) (or via `clasp push` then deploy from the editor). Set “Execute as” to **User accessing the web app** and “Who has access” to **Anyone in your organization** (or your preferred domain scope).
- **Sheet and config:** The dashboard reads from a sheet named **Gold_Testing** in the same spreadsheet as the script. Column layout and constants (e.g. timestamp, email, status, completions) are defined at the top of `Code.js`.

---

## Documentation and resources

- **Notion:** [Incentive Dashboard – Notion](https://www.notion.so/invisibletech/Security-Review-Incentive-Dashboard-Google-Apps-Script-2fc82d3947a680879ab8d393fc944e93?source=copy_link)
