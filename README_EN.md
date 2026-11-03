# Procurement and SAGRILAFT compliance automation

[![Verify code](https://github.com/Dsergio909/sagrilaft-compras-automation/actions/workflows/verificar.yml/badge.svg)](https://github.com/Dsergio909/sagrilaft-compras-automation/actions/workflows/verificar.yml)
![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?logo=googleappsscript&logoColor=white)
![Paid tools](https://img.shields.io/badge/paid_tools-0-2e7d32)

[Español](README.md) · **English**

A system to request, track and audit the documents suppliers must keep up to date under **SAGRILAFT**, Colombia's anti-money-laundering framework, and to flag contracts about to expire. It runs entirely on Google Workspace: Sheets, Gmail, Drive and Calendar, wired together with Apps Script. No Zapier, no Make, no paid APIs.

I built it on my own initiative during my professional internship in a procurement team. This repository holds the code with a sample configuration: it contains no data about the company, its suppliers or its staff.

## What it does

- **Sends requests on its own.** On Tuesdays and Fridays it scans the supplier base and emails everyone within 60 days of their SAGRILAFT date, with the form attached and the document list for companies or individuals.
- **Reads Gmail to see who replied**, through several confidence tiers (below). Nothing uncertain is ever marked as resolved: it leaves a link so a person can check it in one click.
- **Detects bounced emails** and flags the address for correction.
- **Keeps an append-only audit log.** Every action adds a line with date and owner; nothing is overwritten and the sheet is protected.
- **Structures manual follow-up.** Calls, WhatsApp messages and reminders are logged with owner and date; the next follow-up is scheduled in business days (using Colombian public holidays) and stalled cases are flagged automatically.
- **Builds dashboards inside the same sheet:** compliance index, suppliers due for update and a prioritized action plan.
- **Contracts:** a second module sends the team a single email with expired contracts and those about to expire.

## How it detects replies

| Tier | How it searches | Confidence | What it does |
|---|---|---|---|
| 0 | `mailer-daemon` bounce mentioning the supplier's address | High | Marks "bounced" |
| 1 | Thread ID saved when the email was sent | Highest | Updates the status |
| 2 | Exact recipient and subject | High | Updates the status and repairs the link if missing |
| 3a/3b/3c | Sender, company domain or tax ID (NIT) in the body | Medium to low | **Leaves the status alone:** adds a link for human review |

## Architecture

```
Google Sheet
 ├─ BASE DE PROVEEDORES COMPRAS   supplier master data
 ├─ Tabla Control                 status of every request sent
 ├─ Bitácora Auditoría            hidden, protected, append-only
 ├─ DASHBOARD SAGRILAFT           rebuilt from the menu
 ├─ Contratos                     contract matrix
 └─ Control Alertas Contratos     log to avoid repeat alerts

src/ (Apps Script bound to the sheet)
 ├─ Config.gs        emails, IDs, signature and documents: the only file you edit
 ├─ Proveedores.gs   sending, template, detection, dashboard and escalation
 ├─ Contratos.gs     contract alerts and dashboard
 ├─ Menu.gs          sheet menus
 └─ appsscript.json  time zone and V8 runtime
```

## Design decisions

- **When in doubt, a person decides.** A wrong confirmation marks a non-compliant supplier as up to date; an extra review costs a few seconds. That's why heuristic tiers only suggest.
- **Compliance is computed from dates**, not from a status column someone fills in by hand and forgets.
- **History accumulates.** Each system note is appended to the previous ones, with date and a clickable link.
- **Columns by name, not by number.** The dashboard looks up each header; if someone inserts a column, it stops with a clear error instead of silently reading the wrong data.
- **If a dependency fails, work continues.** If the holiday calendar can't be read, the day is treated as a business day rather than blocking follow-up.
- **All configuration in one file.** Changing emails, documents or the signature never touches the logic.

## Setup

1. In a Google Sheet with the tabs above, open **Extensions → Apps Script** and copy the files from `src/`. To paste `appsscript.json`, enable "Show manifest file" in the project settings.
2. Fill in `Config.gs` with your emails, the Drive IDs of the form and signature, and the document list.
3. Reload the sheet and use **🚀 Control Proveedores → ⚙️ Instalación**:
   1. Create the control table and audit log.
   2. Schedule the daily run.
   3. Schedule stalled-case detection.
4. Accept the permissions Google asks for (Sheets, Gmail, Drive and Calendar).

## Verification

`node tools/verificar.js` checks each file's syntax, that no two functions share a name (all Apps Script files share one global scope, so one would silently replace the other) and that every menu item and trigger points to an existing function. It runs on GitHub Actions on every change.

## Limitations

- Gmail returns at most 500 threads per search.
- The system only sees the mailbox of the account running it: if a supplier replies to another address without copying it, there is no way to detect it from code.

## Author

**Sergio García** · International Business, Bogotá · [Portfolio](https://dsergio909.github.io/) · [LinkedIn](https://linkedin.com/in/sergio-david-garcia-celis-836b8a250)
