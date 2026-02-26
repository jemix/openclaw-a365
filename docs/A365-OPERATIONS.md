# A365 Operations & Knowledge Base

This document captures operational knowledge for running OpenClaw A365 with the jemix tenant. It covers the A365 CLI workflow, Graph API permissions, troubleshooting, and lessons learned.

## Table of Contents

- [Environment Overview](#environment-overview)
- [Microsoft Backend: How It All Connects](#microsoft-backend-how-it-all-connects)
- [A365 CLI Workflow](#a365-cli-workflow)
- [Agentic Identity Model](#agentic-identity-model)
- [Inheritable Permissions & Graph Scopes](#inheritable-permissions--graph-scopes)
- [Mail Access: Own vs Shared Mailboxes](#mail-access-own-vs-shared-mailboxes)
- [Token Flow (T1/T2/FIC)](#token-flow-t1t2fic)
- [Frontier License](#frontier-license)
- [Deployment (Mac Mini)](#deployment-mac-mini)
- [Troubleshooting](#troubleshooting)
- [A365 CLI Command Reference](#a365-cli-command-reference)

---

## Environment Overview

| Component | Value |
|-----------|-------|
| Tenant ID | `469d0a07-8515-4814-8ce7-af8d4945938d` |
| Blueprint App ID | `8a09c46a-cd7b-4147-8482-80de28dd72fe` |
| Blueprint Service Principal | `d4dacd17-33fd-4a84-af6d-4fdae0590dd2` |
| Agentic User UPN | `Aila36509e6e6@jemix.com` |
| Agentic User Object ID | `ec1a5658-a46a-4521-a4ee-169952e05dd0` |
| AA_INSTANCE_ID (ServiceIdentity SP) | `d7a4c60a-0cef-4e77-afc4-ee2bd1f1a3a3` |
| Owner | `j.mueller@jemix.com` |
| Owner AAD ID | `dc1c176a-4b0b-4132-bac5-ea84b32a6809` |
| Cloudflare Tunnel | `aipa-x.jemix.com` -> Mac Mini port 3978 |
| GitHub Fork | `jemix/openclaw-a365` (upstream: `SidU/openclaw-a365`) |

## Microsoft Backend: How It All Connects

A365 spans multiple Microsoft portals and systems. Understanding which portal controls what is essential to avoid breaking things.

### The Portal Map

```
┌─────────────────────────────────────────────────────────────────┐
│                     Entra Admin Center                          │
│  (entra.microsoft.com)                                          │
│                                                                 │
│  App Registrations ──── Blueprint app (8a09c46a-...)            │
│  ⚠️  HIDDEN for A365!    DO NOT modify manually!                │
│  The A365 provisioning   Changing API permissions here           │
│  system manages this.    has broken messaging before.            │
│                                                                 │
│  Enterprise Apps ─────── Service Principals                     │
│                          (ServiceIdentity type for instances)   │
│                                                                 │
│  Users ───────────────── Agentic User (Aila36509e6e6@...)       │
│                          Has licenses, mailbox, etc.            │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  Developer Portal (Teams)                        │
│  (dev.teams.microsoft.com)                                      │
│                                                                 │
│  Tools > Agent-Identitaets-Blueprint > Configuration            │
│  └── Notification URL (Messaging Endpoint)                      │
│      Set ONCE on Blueprint level.                               │
│      ⚠️  a365 CLI does NOT set this!                            │
│      Without it: zero messages arrive at your server.           │
│                                                                 │
│  Apps > Your App > Publish                                      │
│  └── Submit manifest to org app catalog                         │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  Teams Admin Center                              │
│  (admin.teams.microsoft.com)                                    │
│                                                                 │
│  Manage Apps ──────────── Published apps & versions             │
│  └── Shows version immediately after publish                    │
│                                                                 │
│  Teams client ─────────── End-user experience                   │
│  └── Version propagation: 5-15 min delay (client cache)         │
│  └── "Instanz erstellen" button creates new Agentic User        │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  Exchange Admin Center                           │
│  (admin.exchange.microsoft.com)                                 │
│                                                                 │
│  Recipients > Mailboxes > Delegation                            │
│  └── Full Access / Send As / Send on Behalf                     │
│      Must be set PER mailbox for the agent's UPN                │
│      Propagation: 15-60 minutes!                                │
│                                                                 │
│  Mail Flow > Accepted Domains                                   │
│  └── Verify all domains are in the same tenant                  │
│      (e.g., jemix.com AND jemix.es in same tenant)              │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  Graph API (beta)                                │
│  (graph.microsoft.com/beta)                                     │
│                                                                 │
│  Inheritable Permissions Endpoint:                              │
│  /applications/{id}/microsoft.graph.agentIdentityBlueprint/     │
│    inheritablePermissions                                       │
│                                                                 │
│  ⚠️  READ: works with az rest                                  │
│  ⚠️  WRITE: blocked by az rest (Directory.AccessAsUser.All)     │
│      → Must use Microsoft Graph PowerShell                      │
│                                                                 │
│  This is where Graph scopes (Mail.ReadWrite.Shared etc.)        │
│  are actually configured. NOT in the Entra UI, NOT via a365 CLI │
└─────────────────────────────────────────────────────────────────┘
```

### Key Relationships

**Blueprint (Entra App Registration) controls:**
- Client ID & Secret (for T1 token)
- Notification/Messaging URL (set via Developer Portal, NOT a365 CLI)
- Inheritable permissions / Graph scopes (set via Graph API)
- Federated Identity Credentials (for T2/FIC token flow)

**Instance (Agentic User in Entra) inherits from Blueprint:**
- All inheritable Graph scopes
- OAuth2 permission grants
- No per-instance configuration needed for permissions

**Exchange (separate from Entra) controls:**
- Mailbox delegation (Full Access, Send As)
- Per-mailbox, per-user grants
- Independent propagation timeline (15-60 min)

### The Two-Layer Security Model

For any resource access, TWO independent layers must both allow it:

```
Layer 1: Graph Scope (Blueprint level)
  "Is the app TYPE of operation allowed?"
  e.g., Mail.ReadWrite.Shared = "shared mailbox access is permitted"

Layer 2: Resource Permission (per resource)
  "Is THIS identity allowed on THIS specific resource?"
  e.g., Exchange Full Access on hola@jemix.es for Aila36509e6e6@jemix.com
```

Missing either layer → 403 Access Denied. This applies to:

| Resource | Layer 1 (Graph Scope) | Layer 2 (Resource Permission) |
|----------|----------------------|-------------------------------|
| Own mailbox | `Mail.ReadWrite` | Automatic (own identity) |
| Shared mailbox | `Mail.ReadWrite.Shared` | Exchange Full Access delegation |
| Own calendar | `Calendars.ReadWrite` | Automatic |
| Shared calendar | `Calendars.ReadWrite.Shared` | Calendar sharing in Outlook |
| SharePoint site | `Sites.Read.All` | Site membership (no .Shared variant) |

### Things That Look Broken But Aren't

1. **"No OAuth2 grants visible"** in Entra Enterprise Apps → Normal for inheritable permissions. They work differently from traditional OAuth2 grants.
2. **App Registration not visible** in Entra → A365 hides it intentionally. Don't try to find/modify it.
3. **Teams shows old version** after publish → Client cache. Admin Center is authoritative.
4. **Token refresh spam in logs** (`getGraphToken called` every second) → Health checks triggering token cache lookups. Harmless at debug level.

## A365 CLI Workflow

### Setup Sequence

The `a365` CLI follows a specific order:

```
1. a365 setup requirements          # Check prerequisites
2. a365 setup infrastructure        # Azure infra (skip with needDeployment: false)
3. a365 setup blueprint             # Create Entra ID app registration
4. a365 setup permissions mcp       # MCP server OAuth2 grants
5. a365 setup permissions bot       # Messaging Bot API grants
```

### Config Files

| File | Purpose | Tracked in git? |
|------|---------|-----------------|
| `a365.config.json` | Static config (tenant, resource group, endpoints) | No (.gitignore) |
| `a365.generated.config.json` | Generated state (IDs, secrets, consent status) | No (.gitignore) |
| `.env` | Runtime env vars for OpenClaw gateway | No (.gitignore) |

**Important:** `a365.config.json` contains `subscriptionId` and `resourceGroup` fields. These are used by `a365 setup infrastructure` for Azure App Service provisioning. With `needDeployment: false` (self-hosting), they are not actively used but still required by the CLI.

### Blueprint vs Instance

- **Blueprint** = The Entra ID app registration. Notification URL, permissions, client secrets, and inheritable scopes are configured here. Shared across all instances.
- **Instance** = An individual Agentic User (e.g., `Aila36509e6e6@jemix.com`). Created via Teams ("Instanz erstellen"). Inherits permissions from the Blueprint.

Key insight: You do NOT set the API/notification URL per instance. It's on the Blueprint level. A new instance can communicate immediately after creation.

### Manifest Publishing

- Use `a365 publish` to push the Teams app manifest
- The `manifest/manifest.json` must NOT contain a `bots` section for API-based agents
- A `bots` section causes Teams to treat the app as a traditional bot -> "Messaging policy block" and "Hinzufuegen" instead of "Instanz erstellen"
- After publishing, it takes ~5-15 minutes for Teams clients to pick up the new version (Admin Center updates immediately)

## Agentic Identity Model

The A365 agent has its own Entra ID user account and acts like a real colleague:

- Users share resources (calendars, mailboxes, SharePoint sites) directly with the agent's UPN
- The agent does NOT use app-level Graph API permissions for resource access
- The T1/T2/FIC token flow produces a delegated token representing the agent's own identity
- Access is scoped to what has been explicitly shared with the agent

**Never manually modify the Blueprint app's API permissions in the Entra UI** - they are managed by the A365 provisioning system. Changing them has broken messaging in the past.

### Role Detection

- When `OWNER_AAD_ID` matches the message sender: `UserRole: Owner`
- All others: `UserRole: Requester`

## Inheritable Permissions & Graph Scopes

### What Are Inheritable Permissions?

Inheritable permissions let agent identities automatically inherit OAuth 2.0 delegated scopes from their parent Blueprint. Newly created instances get these scopes without interactive consent.

### Current Scopes (as of 2026-02-26)

```
Microsoft Graph (00000003-0000-0000-c000-000000000000):
  - Mail.ReadWrite           # Own mailbox read/write
  - Mail.Send                # Send mail as self
  - Mail.ReadWrite.Shared    # Shared/delegated mailbox access
  - Mail.Send.Shared         # Send from shared mailbox
  - Chat.ReadWrite           # Teams chat
  - User.Read.All            # Read user profiles
  - Sites.Read.All           # Read SharePoint sites

Agent 365 Tools (ea9ffc3e-...):
  - McpServersMetadata.Read.All

Messaging Bot API (5a807f24-...):
  - Authorization.ReadWrite
  - user_impersonation

Observability API (9b975845-...):
  - user_impersonation

Power Platform API (8578e004-...):
  - Connectivity.Connections.Read
```

### How to Modify Graph Scopes

The `a365` CLI does NOT manage Graph scopes directly. They must be modified via the Graph API beta endpoint:

```
POST /applications/{blueprintId}/microsoft.graph.agentIdentityBlueprint/inheritablePermissions
```

**Important:** The `az rest` CLI cannot perform write operations on this endpoint because it includes `Directory.AccessAsUser.All` in its token, which Agent APIs block. Use **Microsoft Graph PowerShell** instead:

```powershell
pwsh -Command '
Connect-MgGraph -Scopes "AgentIdentityBlueprint.ReadWrite.All" -NoWelcome

# 1. Delete existing Graph permission entry
Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/applications/8a09c46a-cd7b-4147-8482-80de28dd72fe/microsoft.graph.agentIdentityBlueprint/inheritablePermissions/00000003-0000-0000-c000-000000000000"

# 2. Recreate with updated scopes (include ALL scopes, old + new)
$body = @{
  resourceAppId = "00000003-0000-0000-c000-000000000000"
  inheritableScopes = @{
    "@odata.type" = "microsoft.graph.enumeratedScopes"
    scopes = @(
      "Mail.ReadWrite",
      "Mail.Send",
      "Mail.ReadWrite.Shared",
      "Mail.Send.Shared",
      "Chat.ReadWrite",
      "User.Read.All",
      "Sites.Read.All"
    )
  }
} | ConvertTo-Json -Depth 3

Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/applications/8a09c46a-cd7b-4147-8482-80de28dd72fe/microsoft.graph.agentIdentityBlueprint/inheritablePermissions" -Body $body -ContentType "application/json"
'
```

**Why DELETE + POST?** The Graph API does not support PATCH for inheritable permissions. You must delete the existing entry and recreate it with all scopes (old + new).

### Verifying Permissions

Read-only verification works with `az rest` (no PowerShell needed):

```bash
az rest --method GET --url "https://graph.microsoft.com/beta/applications/8a09c46a-cd7b-4147-8482-80de28dd72fe/microsoft.graph.agentIdentityBlueprint/inheritablePermissions"
```

### Blocked Scopes

Some high-privilege scopes are blocked by platform policy and cannot be made inheritable. If you encounter policy errors, remove the blocked scope. The exact blocklist is not publicly documented.

## Mail Access: Own vs Shared Mailboxes

### Own Mailbox

- Scope: `Mail.ReadWrite` (delegated)
- Works immediately for the agent's own mailbox
- No additional Exchange configuration needed

### Shared / Delegated Mailboxes

Two layers must BOTH be configured:

1. **Graph Scope**: `Mail.ReadWrite.Shared` must be an inheritable permission on the Blueprint
2. **Exchange Delegation**: The agent's UPN must have **Full Access** on the target mailbox

Without either layer, you get 403 "Access is denied".

**Setting Exchange Full Access:**

Via Exchange Admin Center:
> Recipients -> Mailboxes -> target mailbox -> Delegation -> Read and Manage -> Add `Aila36509e6e6@jemix.com`

Via PowerShell:
```powershell
Add-MailboxPermission -Identity "shared@contoso.com" -User "Aila36509e6e6@jemix.com" -AccessRights FullAccess -AutoMapping $false
```

**Propagation:** Exchange delegation can take 15-60 minutes to propagate. Be patient after granting.

### SharePoint Sites

- Scope: `Sites.Read.All` (for reading), `Sites.ReadWrite.All` (for writing)
- Access controlled via SharePoint site membership (add agent as Member/Visitor)
- No `.Shared` variant exists for SharePoint - it's purely site-level permissions

## Token Flow (T1/T2/FIC)

Three-tier token acquisition:

1. **T1 Token**: Client credentials grant using Blueprint App ID + Client Secret + `fmi_path`
2. **T2 Token**: JWT bearer assertion using T1 token
3. **Agent Token**: User FIC (Federated Identity Credential) for the agent's own identity

The `AGENT_IDENTITY` env var (UPN) is critical for the FIC step. A wrong UPN causes `AADSTS50034: user account does not exist`.

Tokens are cached in-memory with a 5-minute buffer before expiration.

## Frontier License

The **Microsoft Agent 365 Frontier** license (~30 EUR/month) includes a surprisingly comprehensive set of services:

| Service | Plan |
|---------|------|
| Exchange | Plan 1 (50 GB mailbox) |
| Teams | Full |
| Phone System | Full |
| Azure AD | P2 Premium |
| Office Apps | Full suite (Outlook, Word, Excel, etc.) |
| Power BI | Full |
| Stream | Plan 2 |
| PowerApps | Full |
| Planner | Full |
| eDiscovery + DLP | Full |
| SharePoint/OneDrive | Included |

**Key takeaway:** The Frontier license provides a full mailbox out of the box. No additional Exchange Online license is needed. A "Business Basic" license is redundant if Frontier is assigned.

## Deployment (Mac Mini)

### Infrastructure

- M4 Mac Mini running macOS
- User: `openclaw` (accessible via `ssh openclaw@192.168.178.133`)
- Cloudflare Tunnel: `aipa-x.jemix.com` -> localhost:3978
- OpenClaw runs plain (no Docker): `pnpm openclaw gateway`

### Plugin Location

```
~/.openclaw/aipa-x/          # Plugin root (git repo)
├── src/                      # Source files
├── node_modules/             # Dependencies
├── .env                      # Runtime config (secrets)
├── a365.generated.config.json # Local state (not tracked)
└── ...
```

### Updating the Plugin

```bash
ssh openclaw@192.168.178.133
cd ~/.openclaw/aipa-x
git pull origin main
# Restart gateway (Ctrl+C existing, then re-run)
pnpm openclaw gateway
```

### Git Remotes (on Mac Mini)

- `origin` -> `jemix/openclaw-a365` (fork)

### Git Remotes (local dev machine)

- `origin` -> `jemix/openclaw-a365` (fork)
- `upstream` -> `SidU/openclaw-a365` (original)

## Troubleshooting

### AADSTS50034: user account does not exist

**Cause:** Wrong `AGENT_IDENTITY` in `.env`. Check the UPN matches the Agentic User exactly.

**Fix:** Update `.env` with correct UPN, restart gateway.

### 403 Access Denied on Shared Mailbox

**Causes (check in order):**
1. Missing `Mail.ReadWrite.Shared` inheritable scope on Blueprint
2. Missing Exchange Full Access delegation for the agent on the target mailbox
3. Propagation delay (wait 15-60 min after granting delegation)
4. Application Access Policy blocking the app (rare)

### EADDRINUSE on Port 3978

**Cause:** Health monitor restarting the provider while server is still running.

**Fix:** The `a365ServerActive` guard in `monitor.ts` prevents this. If it still occurs, check for zombie processes: `lsof -i :3978`

### sendActivity Failed / Typing Indicator Failed

**Cause:** Race condition with double typing indicators. Usually intermittent - messages still get through.

**Severity:** Low. Not a fundamental auth issue.

### Agent APIs do not support Directory.AccessAsUser.All

**Cause:** Using `az rest` for write operations on Agent Identity Blueprint endpoints.

**Fix:** Use Microsoft Graph PowerShell (`Connect-MgGraph -Scopes "AgentIdentityBlueprint.ReadWrite.All"`) instead. Read operations via `az rest` work fine.

### Teams Shows Old App Version

**Cause:** Client-side caching. Teams Admin Center updates immediately, but the Teams client can take 5-15 minutes.

**Fix:** Wait, or clear Teams cache.

## A365 CLI Command Reference

```bash
# Full setup
a365 setup all
a365 setup all --skip-infrastructure    # If Azure infra exists

# Individual steps
a365 setup requirements                 # Check prerequisites
a365 setup infrastructure               # Azure infra (skip for self-hosting)
a365 setup blueprint                    # Create Entra app registration
a365 setup permissions mcp              # MCP server grants
a365 setup permissions bot              # Messaging Bot API grants

# Blueprint management
a365 setup blueprint --endpoint-only    # Just register messaging endpoint
a365 setup blueprint --update-endpoint <url>  # Change endpoint URL

# Publishing
a365 publish                            # Push manifest to Teams

# Querying (via az CLI)
az ad sp list --filter "startswith(displayName,'Aila')"  # Find SPs
az ad user show --id "Aila36509e6e6@jemix.com"           # Check user

# Dry run (any setup command)
a365 setup permissions mcp --dry-run --verbose
```
