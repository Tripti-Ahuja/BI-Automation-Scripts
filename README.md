# BI Automation Scripts

Python scripts for extracting metadata from **Tableau Server/Cloud** and **Power BI** via their REST APIs. Read-only, VDI-safe, and designed to run in corporate environments (IDLE-compatible).

## Scripts

### 1. `Powerbi.py` — Power BI Admin Data Extractor

Interactive CLI tool that connects to Power BI via MSAL device-code auth and exports gateway connections, workspace details, and detailed report/semantic model inventories.

**Available options:**

| # | Category | What You Get |
|---|----------|-------------|
| 1 | Gateway Connections (DSN) | Connection name, type, server, database, credential type, gateway cluster, users |
| 2 | Workspace Details | Overview with owner(s), report/dataset counts, access roles per user |
| 3 | Reports & Semantic Models | Every report and semantic model per workspace — name, created/modified dates, last refresh date, DSN connection name, gateway cluster |
| 4 | Export ALL | Everything above in a single Excel file |

Option 3 also maps each semantic model back to its **gateway DSN connection** and **cluster name**, so you can see which on-premises data source each model is connected to.

### 2. `tableau_server_info.py` — Tableau Server Metadata Explorer

Interactive CLI tool with a menu-driven interface to fetch and export Tableau Server metadata.

**Available options:**

| # | Category | What You Get |
|---|----------|-------------|
| 1 | All Projects | ID, name, description, permissions, parent, owner |
| 2 | All Workbooks | ID, name, project, owner, URL, dates, size, tags |
| 3 | All Views | ID, name, workbook, owner, URL, dates, tags |
| 4 | All Data Sources | ID, name, type, project, owner, dates, tags, certification |
| 5 | Data Source Connections | Drills into every data source to list each connection's type, server, port, database, username |
| 6 | All Flows | ID, name, description, project, owner, dates, tags |
| 7 | Server / Site Summary | Quick count of all object types |
| 8 | Export ALL | Everything above in a single Excel file |

### 3. `connections.py` — Tableau Deep Connection Extractor

Focused script that extracts **all** data source connections, including those embedded inside workbooks (which don't appear as published data sources). Outputs a single Excel file with five sheets:

- **All Data Sources** — published + unique embedded sources merged
- **All Connections** — every connection from published DS and workbooks
- **Published DS Only** — published data sources
- **Workbook Connections** — connections embedded inside workbooks
- **Summary by Type** — connection type breakdown with counts

## Setup

```bash
pip install -r requirements.txt
```

**Dependencies:** `tableauserverclient`, `openpyxl`, `msal`, `requests`

**Optional (recommended):** `pip install truststore` — uses your OS certificate store for SSL, which is critical in corporate/VDI environments.

## Usage

```bash
python Powerbi.py
python tableau_server_info.py
python connections.py
```

Each script will prompt you for connection details:

- **Power BI** (`Powerbi.py`) — paste your Power BI URL (tenant ID is auto-detected), then sign in via device-code flow in your browser
- **Tableau** (`tableau_server_info.py`, `connections.py`) — paste your Tableau URL (base URL and site ID are auto-detected), then enter your Personal Access Token

## Authentication

| Script | Auth Method |
|--------|------------|
| `Powerbi.py` | MSAL device-code flow (sign in via browser, token held in memory) |
| `tableau_server_info.py` | Personal Access Token (PAT) — never saved to disk |
| `connections.py` | Personal Access Token (PAT) — never saved to disk |

## Output

Excel (`.xlsx`) files are saved to the same directory as the script, with a timestamp:

```
PowerBI_ALL_20260409_143022.xlsx
Tableau_ALL_20260409_143022.xlsx
Tableau_Connections_20260409_143022.xlsx
```

Output files are git-ignored since they may contain sensitive metadata (server names, usernames, etc.).

## Security

- **Read-only** — only GET requests; nothing is modified on either server
- **No credentials on disk** — tokens are held in memory only
- **SSL/TLS verification** — enabled by default via `truststore` or `certifi`
- **Session cleanup** — signs out on exit, even if the script crashes
