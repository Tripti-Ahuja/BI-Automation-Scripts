"""
Power BI Admin Data Extractor
Focus: Gateway Connections (DSN) & Workspace Details
Prerequisites: pip install msal requests openpyxl
     (VDI fix): pip install truststore
"""
import sys, os, re, time, json
from datetime import datetime

try:
    import msal, requests
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("\n  Run this first:  pip install msal requests openpyxl\n")
    sys.exit(1)

# ---------------------------------------------------------------------------
# SSL handling for corporate VDI / proxy environments
# ---------------------------------------------------------------------------
_SSL_VERIFY = True
_MSAL_HTTP_CLIENT = None

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    _SSL_VERIFY = False
    _s = requests.Session()
    _s.verify = False
    _MSAL_HTTP_CLIENT = _s

if not _SSL_VERIFY:
    print("  [!] truststore not found — SSL verification disabled.")
    print("      For the proper fix run:  pip install truststore\n")

API = "https://api.powerbi.com/v1.0/myorg"
ADM = API + "/admin"
CID = "ea0616ba-638b-4df5-95b9-636659ae5121"
SCOPES = ["https://analysis.windows.net/powerbi/api/.default"]
MAX_RETRIES = 3
RETRY_DELAY = 5


# ---------------------------------------------------------------------------
# Session — MSAL auth + automatic silent token refresh
# ---------------------------------------------------------------------------
class Session:

    def __init__(self, app, account, token, expires_in):
        self._app = app
        self._account = account
        self._token = token
        self._expiry = time.time() + expires_in

    def token(self):
        if time.time() < self._expiry - 300:
            return self._token
        result = self._app.acquire_token_silent(SCOPES, account=self._account)
        if result and "access_token" in result:
            self._token = result["access_token"]
            self._expiry = time.time() + result.get("expires_in", 3600)
            print("  [Token refreshed]")
        return self._token


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
def login(tenant):
    kwargs = {}
    if _MSAL_HTTP_CLIENT is not None:
        kwargs["http_client"] = _MSAL_HTTP_CLIENT
    app = msal.PublicClientApplication(
        CID, authority="https://login.microsoftonline.com/" + tenant, **kwargs
    )
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("  Login failed:", flow.get("error_description", "Unknown"))
        return None
    print("\n  1. Open:  https://microsoft.com/devicelogin")
    print("  2. Enter code:  " + flow["user_code"])
    print("  3. Sign in with your work email\n")
    print("  Waiting for sign-in...\n")
    r = app.acquire_token_by_device_flow(flow)
    if "access_token" in r:
        print("  Signed in!\n")
        accounts = app.get_accounts()
        return Session(app, accounts[0] if accounts else None,
                       r["access_token"], r.get("expires_in", 3600))
    print("  Sign-in failed:", r.get("error_description", "Unknown"))
    return None


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------
def get(session, url):
    """Paginated GET.  Returns list on success, None on first-page failure."""
    items, first = [], True
    while url:
        h = {"Authorization": "Bearer " + session.token()}
        ok = False
        for attempt in range(MAX_RETRIES):
            try:
                r = requests.get(url, headers=h, timeout=60, verify=_SSL_VERIFY)
            except Exception as e:
                print(f"  [Request error: {e}]")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
                    continue
                return None if first else items
            if r.status_code == 200:
                d = r.json()
                items.extend(d.get("value", []))
                url = d.get("@odata.nextLink")
                ok = True
                break
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", RETRY_DELAY * (attempt + 1)))
                print(f"  [Rate limited — retrying in {wait}s]")
                time.sleep(wait)
                continue
            if r.status_code == 401:
                h = {"Authorization": "Bearer " + session.token()}
                continue
            return None if first else items
        if not ok:
            return None if first else items
        first = False
    return items


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _parse_conn(cd):
    if isinstance(cd, dict):
        return cd
    if isinstance(cd, str) and cd.strip():
        try:
            parsed = json.loads(cd)
            if isinstance(parsed, dict):
                return parsed
        except (json.JSONDecodeError, ValueError):
            return {"raw": cd}
    return {}


def _resolve_type(api_type, cd):
    if api_type != "Extension":
        return api_type
    for key in ("extensionDataSourceKind", "kind", "type"):
        val = cd.get(key, "")
        if val:
            return val
    path = cd.get("extensionDataSourcePath", "")
    if "snowflake" in path.lower():
        return "Snowflake"
    return "Extension"


def _resolve_server(cd, api_type):
    server = cd.get("server", "")
    if server:
        return server
    if api_type == "Extension":
        return cd.get("extensionDataSourcePath", cd.get("path", cd.get("raw", "")))
    return cd.get("path", cd.get("url", cd.get("raw", "")))


_ACCESS_LABELS = {
    "Read": "User",
    "ReadOverrideEffectiveIdentity": "User with resharing",
    "Owner": "Owner",
    "Write": "Owner",
    "None": "None",
}

_ACCESS_FIELD_NAMES = [
    "datasourceAccessRight",
    "datasourceUserAccessRight",
    "accessRight",
    "role",
    "permissions",
    "gatewayDatasourceUserAccessRight",
]


def _best_access_right(user_dict):
    for field in _ACCESS_FIELD_NAMES:
        val = user_dict.get(field)
        if val and str(val) != "None":
            return str(val)
    return user_dict.get("datasourceAccessRight", "")


def _get_ws_users(s, wid):
    """Try multiple API paths to get workspace users."""
    # Path 1: admin per-workspace with $top
    users = get(s, ADM + f"/groups/{wid}/users?$top=500")
    if users:
        return users
    # Path 2: standard API (works if admin has implicit workspace access)
    users = get(s, f"{API}/groups/{wid}/users")
    if users:
        return users
    return []


def _get_ws_item_count(s, wid, item_type):
    """Try multiple API paths to count reports/datasets in a workspace."""
    items = get(s, ADM + f"/groups/{wid}/{item_type}")
    if items is not None:
        return len(items)
    items = get(s, f"{API}/groups/{wid}/{item_type}")
    if items is not None:
        return len(items)
    return 0


# ===================================================================
#  OPTION 1:  GATEWAY CONNECTIONS (DSN)
# ===================================================================
def fetch_gateway_data(s):
    print("  [Getting gateways...]")
    gw_list = get(s, ADM + "/gateways")
    if gw_list is None:
        gw_list = get(s, API + "/gateways") or []

    if not gw_list:
        print("  [No gateways found]")
        return [], []

    print(f"  [Found {len(gw_list)} gateway cluster(s)]\n")

    conn_rows = []
    user_rows = []

    for gw in gw_list:
        gw_id = gw.get("id")
        if not gw_id:
            continue
        gw_cluster = gw.get("name", "")

        sources = get(s, ADM + f"/gateways/{gw_id}/datasources")
        if sources is None:
            sources = get(s, f"{API}/gateways/{gw_id}/datasources") or []
        print(f"  Gateway: {gw_cluster} — {len(sources)} connections")

        for src in sources:
            src_id = src.get("id", "")
            src_name = src.get("datasourceName", "")
            raw_type = src.get("datasourceType", "")
            cd = _parse_conn(src.get("connectionDetails", ""))
            conn_type = _resolve_type(raw_type, cd)
            server = _resolve_server(cd, raw_type)

            # Try admin endpoint first, fall back to standard
            users = get(s, ADM + f"/gateways/{gw_id}/datasources/{src_id}/users")
            if users is None:
                users = get(s, f"{API}/gateways/{gw_id}/datasources/{src_id}/users")
            if users is None:
                users = []

            user_summary = ", ".join(
                u.get("displayName") or u.get("emailAddress", "")
                for u in users
            )

            conn_rows.append({
                "Connection Name": src_name,
                "Connection Type": conn_type,
                "Server": server,
                "Database": cd.get("database", ""),
                "URL": cd.get("url", ""),
                "Path": cd.get("path", cd.get("extensionDataSourcePath", "")),
                "Users": user_summary,
                "# Users": len(users),
                "Gateway Cluster": gw_cluster,
                "Credential Type": src.get("credentialType", ""),
            })

            for u in users:
                raw_right = _best_access_right(u)
                user_rows.append({
                    "Connection Name": src_name,
                    "Connection Type": conn_type,
                    "Gateway Cluster": gw_cluster,
                    "User Name": u.get("displayName", ""),
                    "Email": u.get("emailAddress", ""),
                    "Access Role": _ACCESS_LABELS.get(raw_right, raw_right),
                    "Access (Raw API)": raw_right,
                    "Principal Type": u.get("principalType", ""),
                    "All API Fields": json.dumps(u, default=str),
                })

            time.sleep(0.1)

    print(f"\n  Done — {len(conn_rows)} connections, {len(user_rows)} user entries\n")
    return conn_rows, user_rows


# ===================================================================
#  OPTION 2:  WORKSPACE DETAILS
# ===================================================================
def fetch_workspace_data(s):
    print("  [Getting workspace list...]")
    ws_list = get(s, ADM + "/groups?$top=5000")
    if ws_list is None:
        ws_list = get(s, API + "/groups") or []

    if not ws_list:
        print("  [No workspaces found]")
        return [], []

    total = len(ws_list)
    print(f"  [Found {total} workspaces — fetching users, reports, datasets per workspace]")
    print(f"  [This will take a few minutes...]\n")

    overview_rows = []
    access_rows = []

    for i, ws in enumerate(ws_list):
        wid = ws.get("id", "")
        ws_name = ws.get("name", "")

        if (i + 1) % 25 == 0 or i == 0:
            print(f"  [{i + 1}/{total}] {ws_name[:40]}", flush=True)

        users = _get_ws_users(s, wid)
        rpt_count = _get_ws_item_count(s, wid, "reports")
        ds_count = _get_ws_item_count(s, wid, "datasets")

        owners = [
            u.get("displayName") or u.get("emailAddress", "")
            for u in users
            if u.get("groupUserAccessRight") == "Admin"
        ]

        overview_rows.append({
            "Workspace Name": ws_name,
            "Workspace ID": wid,
            "Type": ws.get("type", ""),
            "State": ws.get("state", ""),
            "Owner(s)": ", ".join(owners),
            "# Reports": rpt_count,
            "# Semantic Models": ds_count,
            "# Users/Groups with Access": len(users),
            "Capacity ID": ws.get("capacityId", ""),
        })

        for u in users:
            access_rows.append({
                "Workspace": ws_name,
                "Workspace ID": wid,
                "User / Group Name": u.get("displayName", ""),
                "Email": u.get("emailAddress", ""),
                "Role": u.get("groupUserAccessRight", ""),
                "Principal Type": u.get("principalType", ""),
                "Identifier (UPN)": u.get("identifier", ""),
            })

        time.sleep(0.15)

    print(f"\n  Done — {len(overview_rows)} workspaces, {len(access_rows)} access entries\n")
    return overview_rows, access_rows


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
_HDR_FONT = Font(bold=True, color="FFFFFF")
_HDR_FILL = PatternFill("solid", fgColor="2F5496")
_HDR_ALIGN = Alignment(horizontal="center")


def _write_sheet(ws, rows):
    if not rows:
        ws["A1"] = "No data"
        return
    heads = list(rows[0].keys())
    for ci, h in enumerate(heads, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.alignment = _HDR_FONT, _HDR_FILL, _HDR_ALIGN
    for ri, row in enumerate(rows, 2):
        for ci, h in enumerate(heads, 1):
            v = row.get(h, "")
            ws.cell(row=ri, column=ci, value=str(v) if isinstance(v, (dict, list)) else v)
    for ci, h in enumerate(heads, 1):
        sample_lens = [len(str(row.get(h, ""))) for row in rows[:200]]
        max_len = max(len(h), max(sample_lens)) if sample_lens else len(h)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 14), 50)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _out_path(name):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        f"PowerBI_{name}_{ts}.xlsx")


def _save_multi(sheets, filename):
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name[:31])
        _write_sheet(ws, rows)
    path = _out_path(filename)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Interactive CLI
# ---------------------------------------------------------------------------
def main():
    print("\n  === POWER BI DATA EXTRACTOR ===\n")
    url = input("  Paste your Power BI URL (or press Enter): ").strip() or "app.powerbi.com"
    m = re.search(r"ctid=([a-f0-9\-]{36})", url, re.I)
    tenant = m.group(1) if m else "common"

    session = login(tenant)
    if not session:
        input("\n  Press Enter to close...")
        return

    while True:
        print("\n  OPTIONS")
        print("  " + "-" * 35)
        print("  1. Gateway Connections (DSN)")
        print("  2. Workspace Details")
        print("  3. Export ALL (both in one file)")
        print("  0. Exit")

        ch = input("\n  Pick (0-3): ").strip()

        if ch == "0":
            print("\n  Bye!\n")
            break

        if ch == "1":
            print("\n  Fetching Gateway Connections...\n")
            conns, users = fetch_gateway_data(session)
            sheets = {"Connections": conns, "Connection Users": users}
            path = _save_multi(sheets, "Gateway_Connections")
            print(f"  Saved -> {path}")
            print(f"  Sheets:  Connections ({len(conns)} rows)  |  "
                  f"Connection Users ({len(users)} rows)")
            input("\n  Press Enter to go back...")

        elif ch == "2":
            print("\n  Fetching Workspace Details...\n")
            overview, access = fetch_workspace_data(session)
            sheets = {"Workspace Overview": overview, "Workspace Access": access}
            path = _save_multi(sheets, "Workspace_Details")
            print(f"  Saved -> {path}")
            print(f"  Sheets:  Workspace Overview ({len(overview)} rows)  |  "
                  f"Workspace Access ({len(access)} rows)")
            input("\n  Press Enter to go back...")

        elif ch == "3":
            print("\n  Fetching everything...\n")
            print("  --- Gateway Connections ---\n")
            conns, conn_users = fetch_gateway_data(session)
            print("  --- Workspace Details ---\n")
            overview, ws_access = fetch_workspace_data(session)

            sheets = {
                "Connections": conns,
                "Connection Users": conn_users,
                "Workspace Overview": overview,
                "Workspace Access": ws_access,
            }
            path = _save_multi(sheets, "ALL")
            total = sum(len(v) for v in sheets.values())
            print(f"  Exported {total} total rows -> {path}")
            print(f"    Connections:        {len(conns)}")
            print(f"    Connection Users:   {len(conn_users)}")
            print(f"    Workspace Overview: {len(overview)}")
            print(f"    Workspace Access:   {len(ws_access)}")
            print("\n  NOTE: file contains sensitive data (server names, emails, etc.)")
            input("\n  Press Enter to go back...")

        else:
            print("  Invalid choice.")


if __name__ == "__main__":
    main()
