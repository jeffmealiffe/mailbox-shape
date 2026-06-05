# mailbox-shape

Analyze the shape of a Microsoft 365 mailbox via the Microsoft Graph API.

## What it reports

- **Folder shape** — full folder tree, item counts, total size per folder, share of items with attachments, breakdown by item type.
- **Message size percentiles** — p50/p75/p90/p95/p99 of message size, split by sent vs. received.
- **Read vs. ignored** — share of received messages that were ever read, bucketed by folder and sender domain.
- **Volume over time** — messages sent, received, and moved into folders, bucketed by day/week/month.

## Auth

Uses MSAL device-code flow against a public-client Azure AD app registration. You'll need to register an application in your tenant (or use a multi-tenant app) and grant delegated `Mail.Read` (and `Mail.ReadBasic` for lighter scans). Set:

```
MAILBOX_SHAPE_CLIENT_ID=<your app client id>
MAILBOX_SHAPE_TENANT_ID=<tenant id or "common" / "consumers">
```

Tokens are cached on disk so you only sign in once per session.

## Install

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
```

## Usage

```powershell
mailbox-shape folders          # folder tree + sizes
mailbox-shape sizes            # message size percentiles
mailbox-shape read-ratio       # read vs. ignored
mailbox-shape volume --by week # volume buckets
mailbox-shape all              # everything, written to ./report/
```

## Status

Early scaffold. See `src/mailbox_shape/analyzers/` for the per-report modules.
