# mailbox-shape

Statistics on the shape, size, and flow of a Microsoft 365 mailbox via the Microsoft Graph API.

Walks your mailbox and reports on folder structure and sizes, message size percentiles, read-vs-ignored ratios, top senders and recipients, volume and rates over time, attachment shares, and item-type mixes. Output goes to the terminal as rich-formatted tables, or to a single self-contained HTML report.

## What it reports

| Command | Reports |
|---|---|
| `folders` | Full folder tree with item counts and PR_MESSAGE_SIZE_EXTENDED-based sizes, recursive by default. |
| `sizes` | p50/p75/p90/p95/p99 message size, sent vs. received. Samples 5k per direction by default. |
| `read-ratio` | Share of received messages ever marked read, grouped by sender domain — strong unsubscribe signal at the bottom of the table. |
| `senders` | Most frequent sender addresses across the Inbox subtree. |
| `recipients` | Most frequent recipient addresses across Sent Items (To + Cc). |
| `volume` | Sent / received / filed counts bucketed by day, week, or month. "Filed" = received messages that ended up in an Inbox subfolder rather than Inbox root. |
| `rates` | Per-day timeline over the last N days, with working-hour-normalized rates that ignore nights and weekends. |
| `attachment-ratio` | Per-folder share of messages with attachments. |
| `message-types` | Per-folder distribution of message subtypes — separates plain mail from `eventMessage`, `eventMessageRequest`, `eventMessageResponse`, etc. |
| `report` | Single self-contained HTML file covering every analyzer above, with inline SVG charts. |
| `whoami` | Diagnostic: what account / tenant / scopes the cached access token is bound to. |
| `raw-folder`, `raw-message` | Diagnostic: dump raw Graph JSON for a folder or a single message. |

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
```

Requires Python 3.11+. On Windows, the `tzdata` PyPI package is pulled automatically so `zoneinfo` works.

## App registration

You need a public-client Azure AD (Entra) app registration with at least delegated `Mail.Read` granted. To target other users' mailboxes via `--mailbox`, also add delegated `Mail.Read.Shared`.

Steps (Azure Portal → **Microsoft Entra ID** → **App registrations** → **New registration**):

1. **Name**: `mailbox-shape` (anything works).
2. **Supported account types**: choose what fits your tenant. Multi-tenant + personal Microsoft accounts is the broadest option.
3. **Redirect URI**: leave blank — device-code flow doesn't use one.
4. After creation, go to **Authentication** → **Advanced settings** → flip **Allow public client flows** to **Yes**, save.
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   - `Mail.Read` (required)
   - `Mail.Read.Shared` (optional, only if you'll use `--mailbox`)
6. Copy the **Application (client) ID** from the **Overview** page.

## Configuration

Copy `.env.example` to `.env` and fill in:

```
MAILBOX_SHAPE_CLIENT_ID=<your app client id GUID>
MAILBOX_SHAPE_TENANT_ID=common
```

`MAILBOX_SHAPE_TENANT_ID` can be `common` (multi-tenant + personal accounts), `organizations` (any work/school tenant), `consumers` (personal Microsoft accounts only), or a specific tenant GUID for single-tenant apps.

Tokens are cached at `~/.mailbox-shape/token_cache.bin` so subsequent runs don't prompt for sign-in.

## Usage

```powershell
# First-run smoke test — triggers device-code sign-in
mailbox-shape whoami

# Folder structure and sizes
mailbox-shape folders                   # recursive (default)
mailbox-shape folders --own             # per-folder own values, no rollup

# Message size percentiles (sent vs received, sampled)
mailbox-shape sizes
mailbox-shape sizes --limit 0           # full scan, slow on large mailboxes

# Sender / recipient / read-ratio
mailbox-shape read-ratio
mailbox-shape senders --top 30
mailbox-shape recipients

# Volume by day / week / month
mailbox-shape volume --by month

# Rates timeline over the last N days
mailbox-shape rates --days 30
mailbox-shape rates --days 90 --work-start 8 --work-end 18

# Attachments and item types
mailbox-shape attachment-ratio
mailbox-shape attachment-ratio --root msgfolderroot   # whole visible mailbox
mailbox-shape message-types

# Full HTML report
mailbox-shape report                                  # writes ./mailbox-report.html
mailbox-shape report -o ~/Desktop/report.html
mailbox-shape report --quick                          # fast preview, skips heavy analyzers
```

## Targeting another mailbox

If you've been granted Exchange delegation or shared-mailbox access, you can run any subcommand against that mailbox via the global `--mailbox` option:

```powershell
mailbox-shape --mailbox alice@example.com folders
mailbox-shape --mailbox alice@example.com report -o alice.html
```

Two prerequisites:

1. **`Mail.Read.Shared` (delegated) on the app registration.** The first run with `--mailbox` triggers a fresh MSAL consent flow asking for the new scope.
2. **The authenticated user has Exchange-level access to the target mailbox.** Either it's a shared mailbox you've been added to, or another user has delegated their mailbox to you in Outlook. Without this, Graph returns `403 ErrorAccessDenied` even with the right scope.

## Performance notes

The fast commands (folders, sizes, recipients, rates, whoami) complete in under a minute. The Inbox-subtree-walking commands (read-ratio, senders, volume, attachment-ratio, message-types) take 5-15 minutes on a 400k-message mailbox because Graph paginates at ~500 messages per request and large folders take a moment to serve each page. Transient `502`/`503`/`504` and throttling `429` responses are retried with exponential backoff and `Retry-After` honoring.

`report` runs a single fused Inbox-subtree walk that populates five analyzers from one stream of messages, so the full report is roughly as slow as a single `volume` run — not the sum of all five.

## Graph quirks worth knowing

| Property | Reality |
|---|---|
| `microsoft.graph.message.size` | Not in the documented OData schema. `$select=size` errors. Read it via the MAPI extended property `PR_MESSAGE_SIZE` (tag `0x0E08`, type `Integer`) expanded inline. |
| `microsoft.graph.mailFolder.sizeInBytes` | Same story. Folder size comes from `PR_MESSAGE_SIZE_EXTENDED` (same tag `0x0E08`, type `Long` — 64-bit so folders >2 GB don't overflow). |
| `@odata.type` on messages | Graph emits this annotation only on subtypes (`eventMessage`, `eventMessageRequest`, etc.). Plain mail items omit it. |
| `Mail.Read` vs. `User.Read` | `Mail.Read` lets you read messages and folders. `/me` itself requires `User.Read`, which is why `whoami` falls back to JWT-claim decoding for identity. |
| `$orderby` on big folders | Sorting + the MAPI expand on a 100k+ folder reliably 504s. Analyzers either drop the orderby or stream pages in natural order. |

## Project layout

```
src/mailbox_shape/
  auth.py                MSAL device-code flow, on-disk token cache
  graph.py               Thin httpx-based Graph client with retries
  cli.py                 Click CLI entry point
  report.py              HTML report orchestrator (Jinja2 + matplotlib)
  templates/report.html  HTML report template
  analyzers/
    folders.py           Folder tree walker with PR_MESSAGE_SIZE_EXTENDED rollups
    sizes.py             Message size percentile sampler
    read_ratio.py        Read vs. ignored by sender domain
    people.py            Top senders / recipients
    volume.py            Sent / received / filed by date bucket
    rates.py             Per-day timeline with working-hour normalization
    attachments.py       Attachment-share per folder
    message_types.py     @odata.type distribution per folder
```

## License

MIT — see [LICENSE](LICENSE).
