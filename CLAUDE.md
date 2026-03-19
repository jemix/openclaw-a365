# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OpenClaw A365 is a Microsoft 365 Agents (A365) channel plugin for OpenClaw. It provides native integration with Microsoft 365 through the Bot Framework and Graph API, allowing agents to:
- Receive/send messages through Microsoft 365 Agents (Teams, Outlook)
- Perform calendar, email, and user operations via Microsoft Graph API
- Authenticate as an agentic identity (its own service account, not on-behalf-of-user)

## Deployment Model

This plugin is deployed as a **tgz archive** installed into OpenClaw via `plugins install` — **not** via Docker. Docker/docker-compose in this repo are legacy and not used in production.

**Production setup (Mac Mini, user `openclaw`):**
- Source repo: `~/projects/a365/` (git clone of `jemix/openclaw-a365`)
- Installed extension: `~/.openclaw/extensions/a365/`
- OpenClaw gateway: runs as LaunchAgent `ai.openclaw.gateway`
- Webhook endpoint: `https://aipa-x.jemix.com/api/messages` → port 3978

**Deploy workflow (Mac Mini):**
```bash
cd ~/projects/a365
git pull
pnpm install       # PATH: /opt/homebrew/opt/node/bin + ~/.npm-global/bin
pnpm build         # compiles TypeScript → dist/ (ignores type errors, see below)
pnpm pack          # creates openclaw-a365-<version>.tgz
openclaw plugins install ./openclaw-a365-<version>.tgz
# then ask Jerry to restart the gateway via launchctl
```

**Important:** Never restart the gateway autonomously. Always ask Jerry — a critical task may be running.

## Build Setup

- `tsconfig.json`: `strict: false`, `noEmitOnError: false` — the codebase has SDK type drift against newer OpenClaw types. The compiled JS works correctly at runtime; type errors are cosmetic. Do not attempt to fix them unless explicitly requested.
- `pnpm build` exits 0 regardless of type errors (uses `tsc || true`)
- Test files (`*.test.ts`) are excluded from the build via tsconfig `exclude`

## Architecture

### Core Components

| File | Purpose |
|------|---------|
| `index.ts` | Plugin entry point — exports `plugin` object that registers with OpenClaw |
| `src/channel.ts` | Channel plugin implementation — registers with OpenClaw, provides capabilities |
| `src/monitor.ts` | Bot Framework webhook listener — receives A365 messages on port 3978 |
| `src/token.ts` | Token management — T1/T2/Agent token flow, in-memory cache |
| `src/graph-tools.ts` | Graph API tools for LLM — calendar, email, attachments, user operations |
| `src/outbound.ts` | Sends messages back to Bot Framework |
| `src/conversation-store.ts` | Persists conversation references for proactive messaging |
| `src/types.ts` | TypeScript types and config schemas |
| `skills/a365/SKILL.md` | Plugin skill — loaded by OpenClaw, tells the agent what it can do |

### Message Flow

1. **Inbound**: Teams/Outlook → A365 Service → POST `/api/messages` → `monitor.ts`
2. **Processing**: Bot activity → `extractMessageMetadata()` → LLM processes with Graph tools
3. **Graph API**: LLM uses tools from `graph-tools.ts` → authenticated via `token.ts`
4. **Outbound**: Response → `outbound.ts` → Bot Framework API → Teams/Outlook

### Key Patterns

**Thread-safe context (`src/graph-tools.ts`):**
Uses `AsyncLocalStorage` for isolating request context. Always wrap request handlers with `runWithGraphToolContext()` — this prevents cross-request data leakage.

**Token acquisition (`src/token.ts`):**
- T1 Token: client credentials + fmi_path
- T2 Token: JWT bearer assertion from T1
- Agent Token: User FIC for agent identity
- All tokens cached in-memory with 5-minute buffer before expiration

**Folder tools use display names, not IDs:**
Graph API mail folder IDs (~120 char Base64) get corrupted by LLMs between tool calls. All folder tools accept `displayName` (e.g. `"Inbox"`, `"Archive"`) and resolve to IDs internally via `resolveMailFolderByName()`. Never change this back to ID-based.

**Proactive messaging (`src/conversation-store.ts`):**
Stores conversation references to `~/.openclaw/a365-conversations.json`. Enables sending messages back to conversations later (cron, async tasks).

## Adding New Tools

When adding a Graph API tool to `src/graph-tools.ts`:

1. **Implement** the async function following the existing pattern (validate userId, call `graphRequest`, return `ToolResult`)
2. **Register** in `createGraphTools()` with `name`, `label`, `description`, `parameters` (TypeBox), `execute`
3. **Check permissions** — verify the required Graph scope is already in the delegated scopes. Current scopes: `Mail.ReadWrite`, `Mail.Send`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`, `Chat.ReadWrite`, `User.Read.All`, `Sites.Read.All`, `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`. New scopes require PowerShell + Graph API changes in Entra (see `docs/A365-OPERATIONS.md`)
4. **Update `skills/a365/SKILL.md`** — add the tool to the tool list so the agent knows it exists
5. **Bump version** in `package.json` and `openclaw.plugin.json`
6. **Deploy** using the workflow above

## Configuration

The plugin reads config from two sources (with env var fallback):

1. **`~/.openclaw/openclaw.json`** → `channels.a365` section — contains all operational config with `${ENV_VAR}` interpolation:
   ```json
   {
     "enabled": true,
     "tenantId": "${A365_TENANT_ID}",
     "appId": "${A365_APP_ID}",
     "appPassword": "${A365_APP_PASSWORD}",
     "webhook": { "port": 3978, "path": "/api/messages" },
     "graph": { "aaInstanceId": "${AA_INSTANCE_ID}", "scope": "https://graph.microsoft.com/.default" },
     "agentIdentity": "${AGENT_IDENTITY}",
     "owner": "${OWNER}",
     "ownerAadId": "${OWNER_AAD_ID}"
   }
   ```
   **Important:** `plugins install` does NOT reset this config. But if it ever gets lost (e.g. manual cleanup), it must be restored manually — the plugin will fail silently without `agentIdentity` and `owner`.

2. **`~/.openclaw/.env`** — env vars used as fallbacks by `token.ts`: `A365_APP_ID`, `A365_APP_PASSWORD`, `A365_TENANT_ID`, `AA_INSTANCE_ID`, `AGENT_IDENTITY`, `OWNER`, `OWNER_AAD_ID`

## Skills

`skills/a365/SKILL.md` is the canonical source of truth for what the agent can do. It is loaded automatically by OpenClaw when the a365 channel is active.

**When to update `SKILL.md`:**
- After adding or removing a tool in `graph-tools.ts`
- After adding or removing a Graph permission

Do NOT maintain a separate TOOLS.md for A365 tool documentation — it will drift. The skill is the single source of truth.

## Key Concepts

**Agentic Identity**: The agent has its own Entra ID user account (`AGENT_IDENTITY`). It only accesses resources explicitly shared with it — this is different from traditional "on-behalf-of" OAuth flows. Users share calendars/resources with the agent like they would with a colleague.

**Role detection**: When `OWNER_AAD_ID` matches the message sender, they get `UserRole: Owner`. Others get `UserRole: Requester`.

**No App Permissions in Entra**: The Blueprint app uses delegated tokens (T1/T2/FIC). Never manually add Application permissions to the Blueprint app in Entra UI — only inheritable delegated scopes via Graph API PowerShell.
