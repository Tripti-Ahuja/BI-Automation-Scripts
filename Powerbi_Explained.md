# Powerbi.py — How It Works (Plain English)

This document explains the entire flow of the `Powerbi.py` script in simple, non-technical terms.

---

## What Does This Script Do?

It connects to your organization's **Power BI account** using Microsoft's official APIs (think of APIs as doorways that let our script ask Power BI for information). It then pulls out metadata (information *about* your Power BI content — not the actual data inside reports) and saves it into a nicely formatted **Excel file**.

---

## Step-by-Step Flow

### 1. Starting Up

When you run the script, it first checks if all the required tools (libraries) are installed. If something is missing, it tells you what to install.

It also checks if your corporate VDI has a special security certificate tool called `truststore`. If not, it turns off strict security checks so the script can still talk to Microsoft's servers (this is safe — it just skips certificate verification).

### 2. Sign In (Authentication)

```
You see:
  1. Open: https://microsoft.com/devicelogin
  2. Enter code: ABCD1234
  3. Sign in with your work email
```

**What's happening behind the scenes:**
- The script asks Microsoft: "I'd like to sign in, give me a one-time code"
- You go to the website, type the code, and sign in with your work email
- Microsoft gives the script a **token** (like a temporary pass/badge) that proves who you are
- This token expires after ~1 hour, but the script automatically renews it in the background

**API used:** Microsoft Login (`login.microsoftonline.com`) via MSAL device-code flow

### 3. The Menu

After signing in, you see 4 options:

```
  1. Gateway Connections (DSN)
  2. Workspace Details
  3. Reports & Semantic Models (detailed)
  4. Export ALL (everything in one file)
```

Each option does something different, explained below.

---

## Option 1: Gateway Connections (DSN)

**Purpose:** Get a list of all data source connections that go through your on-premises gateways (the bridges between your local databases and Power BI cloud).

**Flow:**

```
Step 1 → Ask Power BI: "Give me all gateway clusters"
Step 2 → For each gateway cluster, ask: "What data sources (connections) are configured here?"
Step 3 → For each connection, ask: "Who has access to this connection?"
         (This step runs 8 connections at a time in parallel to speed things up)
Step 4 → Save everything into an Excel file with 2 sheets
```

**APIs called:**

| Step | API Endpoint | What It Returns |
|------|-------------|-----------------|
| 1 | `GET /admin/gateways` | List of all gateway clusters (e.g., "PROD-Gateway", "DEV-Gateway") |
| 2 | `GET /admin/gateways/{gatewayId}/datasources` | All connections on that gateway (name, type, server, database) |
| 3 | `GET /admin/gateways/{gatewayId}/datasources/{datasourceId}/users` | Who has access to each connection and their role |

**Excel output — 2 sheets:**

- **Connections** — One row per connection: name, type (SQL Server, Oracle, etc.), server, database, gateway cluster, how many users
- **Connection Users** — One row per user per connection: name, email, access role, principal type

> **Known limitation:** The API always returns the access role as "Read" even for owners. The "Actual Role" column says "Check Power BI UI" because the API doesn't distinguish owners properly.

---

## Option 2: Workspace Details

**Purpose:** Get an overview of every Power BI workspace — who owns it, how many reports/datasets it has, and who has access.

**Flow:**

```
Step 1 → Ask Power BI: "Give me all workspaces"
Step 2 → For each workspace (8 at a time in parallel):
         a. Ask: "Who are the users of this workspace?"
         b. Ask: "What reports are in this workspace?" (just counts)
         c. Ask: "What datasets are in this workspace?" (just counts)
Step 3 → Save everything into an Excel file with 2 sheets
```

**APIs called:**

| Step | API Endpoint | What It Returns |
|------|-------------|-----------------|
| 1 | `GET /admin/groups?$top=5000` | List of all workspaces (name, ID, state, capacity) |
| 2a | `GET /admin/groups/{workspaceId}/users` | Users and their roles (Admin, Member, Contributor, Viewer) |
| 2b | `GET /admin/groups/{workspaceId}/reports` | List of reports (used for counting) |
| 2c | `GET /admin/groups/{workspaceId}/datasets` | List of datasets (used for counting) |

**Excel output — 2 sheets:**

- **Workspace Overview** — One row per workspace: name, state, owner(s), report count, dataset count, user count, capacity ID
- **Workspace Access** — One row per user per workspace: name, email, role (Admin/Member/Contributor/Viewer)

---

## Option 3: Reports & Semantic Models (detailed)

**Purpose:** Get a detailed list of every report and semantic model (dataset) across your entire organization, including when they were last modified and which gateway connections they use.

**This is the most complex option. It works in 3 phases:**

### Phase 1 — Get all reports and datasets at once

Instead of asking workspace-by-workspace (which would take forever), the script uses two special "tenant-wide" admin APIs that return everything in one go.

| API Endpoint | What It Returns |
|-------------|-----------------|
| `GET /admin/reports?$top=5000` | ALL reports across ALL workspaces, including **created date** and **modified date** |
| `GET /admin/datasets?$top=5000` | ALL datasets across ALL workspaces, including metadata |

> These tenant-wide endpoints return `modifiedDateTime` and `createdDateTime` — the per-workspace endpoints from Option 2 do NOT include these dates. That's why Option 3 uses a different approach.

### Phase 2 — Get last refresh dates

For each dataset that is refreshable (not all are), the script asks: "When was this last refreshed?"

| API Endpoint | What It Returns |
|-------------|-----------------|
| `GET /groups/{workspaceId}/datasets/{datasetId}/refreshes?$top=1` | The most recent refresh: when it started and ended |

This runs in parallel (8 at a time) across all refreshable datasets.

### Phase 3 — Map datasets to gateway connections (DSN)

For datasets that use an on-premises gateway, the script figures out which gateway connection (DSN name) they use.

**How it works:**
1. First, the script builds a "lookup table" from all gateways — mapping each gateway datasource ID to its connection name and cluster
2. Then, for each gateway-connected dataset, it asks: "What data sources does this dataset use?"
3. It matches the datasource IDs to find the connection name

| API Endpoint | What It Returns |
|-------------|-----------------|
| `GET /admin/gateways` | All gateways (for building the lookup) |
| `GET /admin/gateways/{gatewayId}/datasources` | All connections per gateway (for building the lookup) |
| `GET /groups/{workspaceId}/datasets/{datasetId}/datasources` | Which data sources a dataset connects to |

### Final assembly

After all 3 phases, the script combines everything into one flat list and saves it.

**Excel output — 1 sheet:**

- **Reports & Semantic Models** — One row per report/dataset:
  - Workspace Name & ID
  - Item Type (Report or Semantic Model)
  - Name & ID
  - Created Date
  - Modified Date (from the tenant-wide API)
  - Last Refresh Date (datasets only)
  - DSN Connection Name (datasets connected to gateways)
  - Gateway Cluster

---

## Option 4: Export ALL

Runs Options 1 + 2 + 3 back to back and saves everything into a **single Excel file** with 5 sheets:

1. Connections
2. Connection Users
3. Workspace Overview
4. Workspace Access
5. Reports & Semantic Models

---

## How the Script Handles Problems

| Problem | What the Script Does |
|---------|---------------------|
| **API returns an error** | Retries up to 3 times with a 5-second wait between attempts |
| **Rate limiting** (too many requests too fast) | Reads the "Retry-After" header from Power BI and waits that long before trying again |
| **Token expires** (session times out) | Automatically gets a new token in the background without interrupting you |
| **SSL certificate issues** (common on VDI) | Falls back to unverified mode so the script still works |
| **Admin API fails** | Falls back to the standard (non-admin) API as a backup |

---

## API Summary — Every Endpoint Used

| # | Endpoint | Used In | Purpose |
|---|----------|---------|---------|
| 1 | `POST login.microsoftonline.com/.../devicecode` | Sign-in | Get a device code for you to enter on the website |
| 2 | `POST login.microsoftonline.com/.../token` | Sign-in | Exchange the device code for an access token |
| 3 | `GET /admin/gateways` | Options 1, 3 | List all gateway clusters |
| 4 | `GET /admin/gateways/{id}/datasources` | Options 1, 3 | List connections on a gateway |
| 5 | `GET /admin/gateways/{id}/datasources/{id}/users` | Option 1 | List users of a connection |
| 6 | `GET /admin/groups?$top=5000` | Options 2, 3 | List all workspaces |
| 7 | `GET /admin/groups/{id}/users` | Option 2 | List users of a workspace |
| 8 | `GET /admin/groups/{id}/reports` | Option 2 | List reports in a workspace (counts only) |
| 9 | `GET /admin/groups/{id}/datasets` | Option 2 | List datasets in a workspace (counts only) |
| 10 | `GET /admin/reports?$top=5000` | Option 3 | ALL reports tenant-wide (with dates) |
| 11 | `GET /admin/datasets?$top=5000` | Option 3 | ALL datasets tenant-wide (with metadata) |
| 12 | `GET /groups/{id}/datasets/{id}/refreshes?$top=1` | Option 3 | Last refresh date for a dataset |
| 13 | `GET /groups/{id}/datasets/{id}/datasources` | Option 3 | Which gateway connections a dataset uses |

All endpoints start with `https://api.powerbi.com/v1.0/myorg/`.
Endpoints starting with `/admin/` require Power BI Admin permissions.

---

## Performance: Why It Uses Parallel Threads

Imagine you have 500 workspaces and need to ask Power BI about each one. Doing them one-by-one would take forever. Instead, the script uses **8 parallel threads** — like having 8 people making phone calls at the same time instead of 1. This makes the script roughly 8x faster.

Option 3 goes even further by using the tenant-wide APIs (`/admin/reports` and `/admin/datasets`) which return everything in just 2 API calls instead of 1,000+.

---

## What Is the CID?

In the script you'll see:
```
CID = "ea0616ba-638b-4df5-95b9-636659ae5121"
```

This is **Microsoft's official public client ID** for Power BI PowerShell. It's not a secret — it's a well-known identifier that tells Microsoft "this app wants to access Power BI APIs." Think of it as a pre-registered app name that Microsoft recognizes. Every organization that uses Power BI can use this same ID.
