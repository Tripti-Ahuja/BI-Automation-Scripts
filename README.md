# BI Automation Scripts

Python scripts for extracting metadata from Tableau Server / Tableau Cloud via the REST API. Read-only, VDI-safe, and designed to run in corporate environments (IDLE-compatible).

## Scripts

### 1. `tableau_server_info.py` — Server Metadata Explorer

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

### 2. `connections.py` — Deep Connection Extractor

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

**Dependencies:** `tableauserverclient`, `openpyxl`

**Optional (recommended):** `pip install truststore` — uses your OS certificate store for SSL, which is critical in corporate/VDI environments.

## Usage

```bash
python tableau_server_info.py
python connections.py
```

Each script will prompt you for:

1. **Tableau URL** — paste your full browser URL (e.g. `https://us-east-1.online.tableau.com/#/site/mysite/explore`); the base URL and site ID are auto-detected
2. **Personal Access Token (PAT)** — token name and value (in-memory only, never saved to disk)

## Authentication

These scripts use **Personal Access Token (PAT)** authentication only. No passwords are stored or written to disk.

**To create a PAT:**
1. Log into Tableau Cloud / Server in your browser
2. Click your profile icon (top-right) → My Account
3. Under "Personal Access Tokens" click **+ Create**
4. Copy the Token Name and Token Value

## Output

Excel (`.xlsx`) files are saved to the same directory as the script, with a timestamp:

```
Tableau_ALL_20260409_143022.xlsx
Tableau_Connections_20260409_143022.xlsx
```

Output files are git-ignored since they may contain sensitive metadata (server names, usernames, etc.).

## Security

- **Read-only** — only GET requests; nothing is modified on the server
- **No credentials on disk** — PAT is held in memory only
- **SSL/TLS verification** — enabled by default via `truststore` or `certifi`
- **Session cleanup** — signs out on exit, even if the script crashes
