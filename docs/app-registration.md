# Azure AD app registration for mailbox-shape

`mailbox-shape` talks to Microsoft Graph as a **public client** using the OAuth 2.0 device-code flow. That requires you (or your tenant admin) to register an application in Microsoft Entra ID (the new name for Azure AD) once. After that, every sign-in uses the same client ID and is cached locally.

This walkthrough covers:
1. [Picking account types](#1-pick-supported-account-types)
2. [Creating the app registration](#2-create-the-app-registration)
3. [Enabling public-client flows](#3-allow-public-client-flows)
4. [Adding API permissions](#4-add-api-permissions)
5. [Recording the client ID and tenant](#5-get-the-client-id-and-tenant-id)
6. [Configuring your `.env`](#6-configure-your-env)
7. [The first sign-in and consent](#7-first-sign-in)
8. [Adding `Mail.Read.Shared` later](#adding-mailreadshared-later)
9. [Troubleshooting](#troubleshooting)

---

## 1. Pick supported account types

Open https://entra.microsoft.com and sign in. If you have multiple Microsoft accounts cached in the browser, the picker can quietly land you in the wrong tenant — verify the account name in the top-right after sign-in matches the mailbox you'll be analyzing.

Decide which account types your app should support before you click anything. The choice is locked in at creation (changeable later, but it's a hassle).

| Choice | When to use | `MAILBOX_SHAPE_TENANT_ID` value |
|---|---|---|
| **Accounts in this organizational directory only** (single-tenant) | You only need to read mailboxes inside one specific Microsoft 365 tenant and want the tightest blast radius. | The tenant's GUID — find it in **Entra ID** → **Overview**. |
| **Accounts in any organizational directory** (multi-tenant) | You want the same app to work for any work/school tenant. | `organizations` |
| **Accounts in any organizational directory and personal Microsoft accounts** | You want to cover both work/school tenants and personal outlook.com/hotmail.com accounts. Broadest option. | `common` |
| **Personal Microsoft accounts only** | Outlook.com / hotmail.com / personal MSA only. | `consumers` |

The `common`, `organizations`, and `consumers` values are special MSAL endpoints; you don't need to know your own tenant GUID to use them.

For a personal project against your own mailbox, **"Accounts in any organizational directory and personal Microsoft accounts"** + `MAILBOX_SHAPE_TENANT_ID=common` works for any case and survives moving the mailbox between tenants.

---

## 2. Create the app registration

In the Entra portal:

1. Left nav → **Applications** → **App registrations** → **+ New registration** (top of the page).
2. Fill in the form:
   - **Name**: anything that helps you recognize it later. `mailbox-shape` is fine.
   - **Supported account types**: the choice from [step 1](#1-pick-supported-account-types).
   - **Redirect URI**: leave blank. Device-code flow doesn't use a redirect URI; trying to set one as a public client may trigger validation errors.
3. Click **Register**. You land on the **Overview** page for the new app. Keep this tab open.

---

## 3. Allow public client flows

Device-code flow is a "public client" pattern — the client has no secret, just a client ID. Entra blocks this by default; you have to opt in.

1. In your app's left nav, click **Authentication**.
2. Scroll to the very bottom — past the Platform Configurations and Implicit Grant sections — to **Advanced settings**.
3. Find **"Allow public client flows"** and flip it to **Yes**.
4. Click **Save** at the top of the page.

If you skip this step, sign-in fails with **AADSTS7000218** ("The request body must contain the following parameter: 'client_assertion' or 'client_secret'.").

---

## 4. Add API permissions

mailbox-shape needs at minimum **`Mail.Read`** (delegated). Optionally also add `Mail.Read.Shared` if you want to use `--mailbox` to read another user's delegated mailbox.

1. In your app's left nav, click **API permissions**.
2. Click **+ Add a permission**.
3. Select **Microsoft Graph**.
4. Select **Delegated permissions** (NOT application permissions — those are for server-side daemons with admin consent and don't apply here).
5. In the search box, type `Mail.Read`. Check:
   - **`Mail.Read`** — Read user mail. **Required.**
   - **`Mail.Read.Shared`** *(optional)* — Read user and shared mail. Only needed if you plan to use `--mailbox` to read someone else's mailbox you've been granted access to.
6. Click **Add permissions**.

You should now see `Mail.Read` (and `Mail.Read.Shared` if you added it) listed under "Configured permissions" with **Type = Delegated** and **Status** blank.

**You don't need to click "Grant admin consent"** for these scopes — they're user-consentable, so you'll just be prompted to accept on first sign-in.

> **Note about `User.Read`**: by default, Entra adds `User.Read` to new app registrations. mailbox-shape doesn't strictly need it — it decodes the JWT token claims directly instead of calling `/me`. You can leave `User.Read` granted (it doesn't hurt) or remove it. If you remove it, the `mailbox-shape whoami` command will show `403 ErrorAccessDenied` next to the `/me` row, but the rest of the table still works because it reads from the token's JWT claims.

---

## 5. Get the client ID and tenant ID

On your app's **Overview** page:

- **Application (client) ID** — a GUID like `3cb320b1-cb9a-4bdf-9ffd-dcd2f23099b1`. This is `MAILBOX_SHAPE_CLIENT_ID`.
- **Directory (tenant) ID** — also a GUID. Only relevant if you picked single-tenant in [step 1](#1-pick-supported-account-types). For multi-tenant and personal MSA, use one of the special values (`common`, `organizations`, `consumers`) instead.

Copy these somewhere safe. They're not secrets — public clients have no secrets by definition — but they're tedious to look up again.

---

## 6. Configure your `.env`

In the repo root, copy `.env.example` to `.env` and fill in the values:

```ini
MAILBOX_SHAPE_CLIENT_ID=3cb320b1-cb9a-4bdf-9ffd-dcd2f23099b1
MAILBOX_SHAPE_TENANT_ID=common
```

`.env` is gitignored, but the client ID isn't actually secret. If you commit it by accident the worst case is someone else can attempt sign-in flows against your app registration — they still need their own Microsoft account, the app registration is bound to your tenant settings, and you can revoke the app at any time.

---

## 7. First sign-in

Run any command — `whoami` is the lightest:

```powershell
mailbox-shape whoami
```

The first run will print something like:

```
To sign in, use a web browser to open the page https://login.microsoft.com/device
and enter the code BX58GBG95 to authenticate.
```

Steps:

1. Open the URL on any device (your phone is fine).
2. Enter the code.
3. Sign in with the Microsoft account that owns the mailbox you want to analyze.
4. On the consent screen, accept the requested permissions (`Mail.Read` at minimum). If you're signed into the same browser as multiple Microsoft accounts, double-check that the one being asked to consent matches the one whose mailbox you want.

After consent, the CLI continues, fetches the token, caches it at `~/.mailbox-shape/token_cache.bin`, and runs the command. Subsequent runs skip the device-code prompt entirely until the refresh token expires (~90 days).

---

## Adding `Mail.Read.Shared` later

If you initially set up with only `Mail.Read` and now want to use `--mailbox`:

1. Go back to your app registration → **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**, add `Mail.Read.Shared`, **Add permissions**.
2. Delete the local token cache so MSAL re-requests the new scope:
   ```powershell
   Remove-Item ~/.mailbox-shape/token_cache.bin
   ```
3. Run any `mailbox-shape --mailbox <upn>` command. MSAL will trigger a new device-code flow and ask you to consent to the additional scope.

You can also leave the cache alone and just run something that uses `--mailbox` — MSAL detects the new scope requirement at acquire time and prompts for re-consent, then caches a new token covering both scopes. The cache-delete approach is slightly more deterministic if anything seems stuck.

---

## Troubleshooting

### `AADSTS7000218` — "client is not allowed to use public client flow"

You skipped [step 3](#3-allow-public-client-flows). Go back to **Authentication** → **Allow public client flows** → flip to **Yes** → **Save**.

### `AADSTS65001` — "user or admin has not consented"

The consent screen was dismissed without accepting, or your tenant requires admin consent for the scope. Most likely fixes:

- Retry the sign-in and click **Accept** on the consent screen this time.
- If your tenant requires admin consent for `Mail.Read.Shared` specifically, ask your tenant admin to click **Grant admin consent for `<your tenant>`** on the API permissions page of your app registration.

### `AADSTS50020` — "user account from identity provider does not exist in tenant"

The Microsoft account you signed in with isn't a member of the tenant your app registration is in (or your app is single-tenant and you tried signing in with a personal account). Either:

- Sign out of all Microsoft accounts in your browser and sign in with the right one before entering the device code.
- Change your app's supported account types in **Authentication** → **Supported account types** to a broader option (multi-tenant or include personal MS accounts).

### `403 ErrorAccessDenied` when using `--mailbox`

The scope is fine, but the authenticated user doesn't have Exchange-level access to the target mailbox. Two layers of access matter here:

1. **Graph scope** — `Mail.Read.Shared` granted in the app registration *and* consented to.
2. **Exchange delegation** — in Exchange/Outlook, the target mailbox must explicitly grant the authenticated user "Folder Visible" + "Reviewer" (or higher) on the folders you're trying to read. Either configure that in Outlook ("Folder Permissions" on the target mailbox) or have your Exchange admin set it via `Add-MailboxFolderPermission` in PowerShell.

Graph reports the second-layer failure as `ErrorAccessDenied` — the scope is fine but Exchange said no.

### `ZoneInfoNotFoundError` on Windows

The `tzdata` PyPI package isn't installed. It's a platform-conditional dependency declared in `pyproject.toml`, so:

```powershell
pip install -e .
```

should pull it in. If you're using a manually-pinned environment, `pip install tzdata` explicitly.

### Sign-in prompt loops or token cache misbehaves

```powershell
Remove-Item ~/.mailbox-shape/token_cache.bin
```

Then re-run. The next call triggers a clean device-code flow.
