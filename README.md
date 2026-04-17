# BI Automation Scripts

Python scripts to extract metadata from **Power BI** and **Tableau Server/Cloud** via REST APIs. Read-only, VDI-safe, no credentials saved to disk.

## Scripts

| Script | Platform | What It Does |
|--------|----------|-------------|
| `Powerbi.py` | Power BI | Gateway connections (DSN), workspace details, reports & semantic models with dates and DSN mapping |
| `tableau_server_info.py` | Tableau | Projects, workbooks, views, data sources, connections, flows, site summary |
| `connections.py` | Tableau | Deep connection extractor — includes connections embedded inside workbooks |

## Quick Start

```bash
pip install -r requirements.txt
pip install truststore          # recommended for VDI/corporate environments

python Powerbi.py               # Power BI — sign in via browser device-code
python tableau_server_info.py   # Tableau — enter your Personal Access Token
python connections.py           # Tableau — enter your Personal Access Token
```

## Output

Excel files saved in the same folder with a timestamp (e.g., `PowerBI_ALL_20260409_143022.xlsx`). Output files are git-ignored.

## Security

- **Read-only** — only GET requests, nothing is modified
- **No credentials on disk** — tokens held in memory only
- **SSL/TLS** — uses OS certificate store via `truststore` when available

For a detailed walkthrough of `Powerbi.py` (APIs, flow, options), see [Powerbi_Explained.md](Powerbi_Explained.md).
