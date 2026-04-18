# Powerbi.py — API Reference

All API endpoints used in the script, with line numbers and plain-English descriptions.

All endpoints (3–13) start with `https://api.powerbi.com/v1.0/myorg/`.
Endpoints starting with `/admin/` require Power BI Admin permissions.

| # | API | What It Does | API Endpoint | Line(s) |
|---|-----|-------------|-------------|---------|
| 1 | **Login (Device Code)** | Gives you a one-time code to sign in | `POST login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode` | 84 |
| 2 | **Login (Token)** | Exchanges the code for a temporary pass | `POST login.microsoftonline.com/{tenant}/oauth2/v2.0/token` | 92 |
| 3 | **Get All Gateways** | Lists all gateway clusters | `GET /admin/gateways` | 215, 399 |
| 4 | **Get Gateway Connections** | Lists connections on a gateway | `GET /admin/gateways/{gatewayId}/datasources` | 231, 407 |
| 5 | **Get Connection Users** | Lists who has access to a connection | `GET /admin/gateways/{gatewayId}/datasources/{datasourceId}/users` | 257 |
| 6 | **Get All Workspaces** | Lists every workspace | `GET /admin/groups?$top=5000` | 319, 436 |
| 7 | **Get Workspace Users** | Lists users and roles in a workspace | `GET /admin/groups/{workspaceId}/users` | 201 |
| 8 | **Get Workspace Reports** | Lists reports in a workspace (counts) | `GET /admin/groups/{workspaceId}/reports` | 191 (via _get_items at 342) |
| 9 | **Get Workspace Datasets** | Lists datasets in a workspace (counts) | `GET /admin/groups/{workspaceId}/datasets` | 191 (via _get_items at 343) |
| 10 | **Get All Reports (Tenant-wide)** | Every report across ALL workspaces with dates | `GET /admin/reports?$top=5000` | 448 |
| 11 | **Get All Datasets (Tenant-wide)** | Every dataset across ALL workspaces | `GET /admin/datasets?$top=5000` | 453 |
| 12 | **Get Dataset Refresh History** | Last refresh date for a dataset | `GET /groups/{workspaceId}/datasets/{datasetId}/refreshes?$top=1` | 479 |
| 13 | **Get Dataset Data Sources** | Which gateway connection a dataset uses | `GET /groups/{workspaceId}/datasets/{datasetId}/datasources` | 518 |

## Which Option Uses Which APIs

- **Option 1** (Gateway Connections) — APIs 3, 4, 5
- **Option 2** (Workspace Details) — APIs 6, 7, 8, 9
- **Option 3** (Reports & Semantic Models) — APIs 3, 4, 6, 10, 11, 12, 13
- **Option 4** (Export ALL) — All of the above

## Notes

- APIs 12 and 13 use the standard (non-admin) endpoint because the admin version doesn't support refresh history or dataset datasource lookups.
- APIs 1 and 2 are handled automatically by the MSAL library — the script doesn't call these URLs directly.
- The token from API 2 is valid for 3600 seconds (1 hour) and is auto-refreshed silently.
