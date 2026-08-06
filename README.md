# Serfis Document Management — SharePoint In-Place Editing Demos

Two Blazor Server (.NET 8) applications demonstrating the two proposed architectures for
replacing the download → edit → re-upload workflow with in-place Office editing via
SharePoint Online and Microsoft Graph. Companion design documents:

- *Serfis-SharePoint-Document-Editing-Design.docx* — Scenario 1
- *Serfis-SharePoint-System-of-Record-Scenario2.docx* — Scenario 2 (recommended)

| | Scenario 1 — `Scenario1.TransientEditing` | Scenario 2 — `Scenario2.SharePointMaster` |
|---|---|---|
| System of record | Network share (simulated by a local folder) | SharePoint document library |
| SharePoint role | Temporary editing workspace | Permanent, **sealed** storage |
| On View/Edit | Upload to SharePoint (Track Changes injected), open in Word | JIT permission grant on that one document, open in Word |
| When editing ends (inferred: idle + lock probe) | Sync file back to share, delete from SharePoint | Revoke the JIT grants |
| Extras shown | Live activity feed, manual sweep, co-authoring join | Version history + restore, access DB, reconciliation sweep |
| URL | https://localhost:7191 | https://localhost:7292 |

Both apps show a **live activity feed** so an audience can watch uploads, grants,
lock probes, sync-backs and revocations happen in real time.

## One-time setup

### 1. Entra ID app registration

Create one app registration (both apps can share it):

- **Redirect URIs** (Web): `https://localhost:7191/signin-oidc` and `https://localhost:7292/signin-oidc`
- **Front-channel logout / post-logout**: optional for the demo
- **Client secret**: create one; store it with `dotnet user-secrets` (below)
- **API permissions**:
  - Delegated: `openid`, `profile`, `offline_access`, `User.Read`, `Sites.Selected` (Scenario 1 edits as the user, limited to sites granted to the app)
  - Application: `Sites.Selected` (Scenario 2 custodian + Scenario 1 background sweep)
  - Grant **admin consent**
- **Sites.Selected step 2** (grants nothing until you do this): give the app *write* on the target site:

  ```
  POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
  {
    "roles": ["write"],
    "grantedToIdentities": [
      { "application": { "id": "<app client id>", "displayName": "Serfis Demo" } }
    ]
  }
  ```

### 2. Find the site & drive IDs

```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-path}
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

The default document library's `id` is the **DriveId**.

### 3. Configure each app

In `appsettings.json` set `AzureAd:TenantId`, `AzureAd:ClientId`, `SharePoint:SiteId`,
`SharePoint:DriveId`. Keep the secret out of the file:

```bash
cd Scenario1.TransientEditing   # and again in Scenario2.SharePointMaster
dotnet user-secrets init
dotnet user-secrets set "AzureAd:ClientSecret" "<secret>"
```

### 4. Trust the dev certificate & run

```bash
dotnet dev-certs https --trust
dotnet run --project Scenario1.TransientEditing   # https://localhost:7191
dotnet run --project Scenario2.SharePointMaster   # https://localhost:7292
```

Desktop Word editing (`ms-word:ofe|u|…`) requires desktop Office signed into an
M365 account with access; Office Online editing opens in a new browser tab.

## Demo notes

- **Idle threshold / sweep interval** are in `appsettings.json` (`Demo` section),
  defaulted to 120 s / 30 s so a live audience sees the automatic sync-back or
  revocation within ~2–3 minutes of closing Word. Production values would be larger.
- **Change notifications vs polling**: production subscribes to Microsoft Graph change
  notifications (webhooks from SharePoint) to track last-activity; the demos poll
  `lastModifiedDateTime` because localhost can't receive webhooks. The inference logic
  (inactivity → checkout lock probe → act) is identical.
- **Track Changes** is enforced by injecting `<w:trackChanges/>` (SDK class
  `TrackRevisions`) into `word/settings.xml` before any document reaches SharePoint.
- **Scenario 1 without app-only credentials**: the background sweep disables itself and
  the UI's "Sync back now / Run sweep" buttons still work (they run as the signed-in user).
- **Scenario 2 requires** the app-only credentials — the sealed library has no user
  access by design, so only the service principal can create documents and grant access.
- App state lives in `App_Data/*.json` (and Scenario 1's simulated share in
  `NetworkShare/`); both are git-ignored. Delete them to reset a demo.
