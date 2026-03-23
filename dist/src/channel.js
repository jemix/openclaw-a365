import { DEFAULT_ACCOUNT_ID } from "openclaw/plugin-sdk";
import { resolveA365Credentials, resolveGraphTokenConfig, resolveTokenCallbackConfig, getGraphToken } from "./token.js";
import { createGraphTools } from "./graph-tools.js";
import { a365Outbound, normalizeA365MessagingTarget } from "./outbound.js";
/**
 * Resolve the effective A365Config for a specific account.
 *
 * Multi-account mode: merges top-level defaults with per-account overrides.
 * Single-account mode (no `accounts` key): returns the flat config as-is.
 */
export function resolveAccountA365Config(a365Cfg, accountId) {
    if (!a365Cfg)
        return undefined;
    const accounts = a365Cfg.accounts;
    if (!accounts || !accountId || accountId === DEFAULT_ACCOUNT_ID && !accounts[DEFAULT_ACCOUNT_ID]) {
        // Single-account mode or default account without explicit entry
        return a365Cfg;
    }
    const acct = accounts[accountId];
    if (!acct)
        return undefined;
    // Merge: account-specific fields override top-level defaults.
    // webhook is shared (lives on A365Config only).
    return {
        ...a365Cfg,
        ...acct,
        graph: { ...a365Cfg.graph, ...acct.graph },
        tokenCallback: acct.tokenCallback ?? a365Cfg.tokenCallback,
        businessHours: acct.businessHours ?? a365Cfg.businessHours,
        allowFrom: acct.allowFrom ?? a365Cfg.allowFrom,
        groupAllowFrom: acct.groupAllowFrom ?? a365Cfg.groupAllowFrom,
    };
}
/**
 * List all configured account IDs from an A365Config.
 */
function listA365AccountIds(a365Cfg) {
    if (a365Cfg?.accounts && Object.keys(a365Cfg.accounts).length > 0) {
        return Object.keys(a365Cfg.accounts);
    }
    return [DEFAULT_ACCOUNT_ID];
}
/**
 * Check if Graph API tools can be enabled.
 * Requires either T1/T2 flow config or external token callback.
 */
function isGraphConfigured(cfg) {
    return Boolean(resolveGraphTokenConfig(cfg) || resolveTokenCallbackConfig(cfg));
}
const meta = {
    id: "a365",
    label: "Microsoft 365 Agents",
    selectionLabel: "Microsoft 365 Agents (A365)",
    docsPath: "/channels/a365",
    docsLabel: "a365",
    blurb: "Native A365 channel with Graph API tools for calendar and email.",
    aliases: ["m365agents", "agents365"],
    order: 55,
};
/**
 * Probe A365 configuration to check if it's working.
 */
async function probeA365(cfg) {
    const creds = resolveA365Credentials(cfg);
    if (!creds) {
        return { ok: false, error: "Bot Framework credentials not configured" };
    }
    // Check if Graph API is configured
    const graphConfigured = isGraphConfigured(cfg);
    let graphConnected = false;
    if (graphConfigured && cfg?.agentIdentity) {
        try {
            const token = await getGraphToken(cfg, cfg.agentIdentity);
            graphConnected = Boolean(token);
        }
        catch {
            graphConnected = false;
        }
    }
    return {
        ok: true,
        botId: creds.appId,
        graphConnected,
        owner: cfg?.owner,
    };
}
/**
 * A365 Channel Plugin for OpenClaw.
 *
 * This channel enables:
 * - Receiving messages from Microsoft 365 Agents via Bot Framework
 * - Sending responses back through Bot Framework
 * - Native Graph API tools for calendar, email, and user operations
 *
 * Supports multiple accounts via the `accounts` config map.
 */
export const a365Plugin = {
    id: "a365",
    meta: {
        ...meta,
    },
    pairing: {
        idLabel: "a365UserId",
        normalizeAllowEntry: (entry) => entry.replace(/^(a365|user):/i, ""),
    },
    capabilities: {
        chatTypes: ["direct", "channel", "thread"],
        threads: true,
        media: true,
    },
    agentPrompt: {
        messageToolHints: ({ cfg }) => {
            const a365Cfg = cfg?.channels?.a365;
            const timezone = a365Cfg?.businessHours?.timezone || "America/Los_Angeles";
            const now = new Date();
            const formatter = new Intl.DateTimeFormat("en-US", {
                timeZone: timezone,
                year: "numeric",
                month: "2-digit",
                day: "2-digit",
                hour: "2-digit",
                minute: "2-digit",
                hour12: false,
            });
            const currentDateTime = formatter.format(now);
            const dateOnly = new Date().toLocaleDateString("en-CA", { timeZone: timezone }); // YYYY-MM-DD format
            const hints = [
                "- A365 channel supports direct messages and channel conversations.",
                "- Use Graph API tools (get_calendar_events, create_calendar_event, etc.) for calendar operations.",
                `- Current date/time: ${currentDateTime} (${timezone}). Today's date in ISO format: ${dateOnly}.`,
            ];
            if (a365Cfg?.owner) {
                hints.push(`- Default calendar owner: ${a365Cfg.owner}`);
            }
            const klipyKey = a365Cfg?.klipyApiKey || process.env.KLIPY_API_KEY;
            if (klipyKey) {
                hints.push("- You have a send_gif tool to send animated GIFs inline. Use it sparingly and only when it genuinely fits the moment — celebrations, humor, greetings, empathy. Do NOT use it on every message. Never use it when delivering bad news or discussing serious matters.");
            }
            return hints;
        },
    },
    threading: {
        buildToolContext: ({ context, hasRepliedRef }) => ({
            currentChannelId: context.To?.trim() || undefined,
            currentThreadTs: context.ReplyToId,
            hasRepliedRef,
        }),
    },
    reload: { configPrefixes: ["channels.a365"] },
    config: {
        listAccountIds: (cfg) => {
            const a365Cfg = cfg.channels?.a365;
            return listA365AccountIds(a365Cfg);
        },
        resolveAccount: (cfg, accountId) => {
            const a365Cfg = cfg.channels?.a365;
            const effectiveId = accountId || DEFAULT_ACCOUNT_ID;
            const acctCfg = resolveAccountA365Config(a365Cfg, effectiveId);
            return {
                accountId: effectiveId,
                enabled: acctCfg?.enabled !== false,
                configured: Boolean(resolveA365Credentials(acctCfg)),
                owner: acctCfg?.owner,
            };
        },
        defaultAccountId: (cfg) => {
            const a365Cfg = cfg.channels?.a365;
            const ids = listA365AccountIds(a365Cfg);
            return ids[0] ?? DEFAULT_ACCOUNT_ID;
        },
        setAccountEnabled: ({ cfg, accountId, enabled }) => {
            const a365Cfg = (cfg.channels?.a365 ?? {});
            const accounts = a365Cfg.accounts;
            // Multi-account: update the specific account entry
            if (accounts && accountId && accountId !== DEFAULT_ACCOUNT_ID) {
                return {
                    ...cfg,
                    channels: {
                        ...cfg.channels,
                        a365: {
                            ...a365Cfg,
                            accounts: {
                                ...accounts,
                                [accountId]: { ...accounts[accountId], enabled },
                            },
                        },
                    },
                };
            }
            // Single-account / default
            return {
                ...cfg,
                channels: {
                    ...cfg.channels,
                    a365: { ...a365Cfg, enabled },
                },
            };
        },
        deleteAccount: ({ cfg, accountId }) => {
            const a365Cfg = cfg.channels?.a365;
            const accounts = a365Cfg?.accounts;
            // Multi-account: remove the specific account entry
            if (accounts && accountId && accounts[accountId]) {
                const nextAccounts = { ...accounts };
                delete nextAccounts[accountId];
                const nextA365 = { ...a365Cfg };
                if (Object.keys(nextAccounts).length > 0) {
                    nextA365.accounts = nextAccounts;
                }
                else {
                    delete nextA365.accounts;
                }
                return { ...cfg, channels: { ...cfg.channels, a365: nextA365 } };
            }
            // Single-account: remove entire a365 config
            const next = { ...cfg };
            const nextChannels = { ...cfg.channels };
            delete nextChannels.a365;
            if (Object.keys(nextChannels).length > 0) {
                next.channels = nextChannels;
            }
            else {
                delete next.channels;
            }
            return next;
        },
        isConfigured: (_account, cfg) => {
            const a365Cfg = cfg.channels?.a365;
            const acctCfg = resolveAccountA365Config(a365Cfg, _account.accountId);
            return Boolean(resolveA365Credentials(acctCfg));
        },
        describeAccount: (account) => ({
            accountId: account.accountId,
            enabled: account.enabled,
            configured: account.configured,
        }),
        resolveAllowFrom: ({ cfg, accountId }) => {
            const a365Cfg = cfg.channels?.a365;
            const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
            return acctCfg?.allowFrom?.map(String) ?? [];
        },
        formatAllowFrom: ({ allowFrom }) => allowFrom
            .map((entry) => String(entry).trim())
            .filter(Boolean)
            .map((entry) => entry.toLowerCase()),
    },
    security: {
        resolveDmPolicy: ({ cfg, accountId }) => {
            const a365Cfg = cfg.channels?.a365;
            const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
            return {
                policy: acctCfg?.dmPolicy ?? "pairing",
                allowFrom: acctCfg?.allowFrom ?? [],
                allowFromPath: "channels.a365.",
                approveHint: "Add user ID to channels.a365.allowFrom",
            };
        },
        collectWarnings: ({ cfg, accountId }) => {
            const a365Cfg = cfg.channels?.a365;
            const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
            const groupPolicy = acctCfg?.groupPolicy ?? "allowlist";
            if (groupPolicy !== "open") {
                return [];
            }
            return [
                `- A365 groups: groupPolicy="open" allows any member to trigger. Set channels.a365.groupPolicy="allowlist" + channels.a365.groupAllowFrom to restrict senders.`,
            ];
        },
    },
    setup: {
        resolveAccountId: ({ cfg, accountId }) => {
            if (accountId)
                return accountId;
            const a365Cfg = cfg.channels?.a365;
            const ids = listA365AccountIds(a365Cfg);
            return ids[0] ?? DEFAULT_ACCOUNT_ID;
        },
        applyAccountConfig: ({ cfg, accountId }) => {
            const a365Cfg = (cfg.channels?.a365 ?? {});
            const accounts = a365Cfg.accounts;
            // Multi-account: enable the specific account
            if (accounts && accountId && accountId !== DEFAULT_ACCOUNT_ID) {
                return {
                    ...cfg,
                    channels: {
                        ...cfg.channels,
                        a365: {
                            ...a365Cfg,
                            accounts: {
                                ...accounts,
                                [accountId]: { ...accounts[accountId], enabled: true },
                            },
                        },
                    },
                };
            }
            return {
                ...cfg,
                channels: {
                    ...cfg.channels,
                    a365: { ...a365Cfg, enabled: true },
                },
            };
        },
    },
    messaging: {
        normalizeTarget: normalizeA365MessagingTarget,
        targetResolver: {
            looksLikeId: (raw) => {
                const trimmed = raw.trim();
                if (!trimmed) {
                    return false;
                }
                if (/^conversation:/i.test(trimmed)) {
                    return true;
                }
                if (/^user:/i.test(trimmed)) {
                    const id = trimmed.slice("user:".length).trim();
                    return /^[0-9a-fA-F-]{16,}$/.test(id);
                }
                return trimmed.includes("@thread") || trimmed.includes(":");
            },
            hint: "<conversationId|user:ID|conversation:ID>",
        },
    },
    directory: {
        self: async () => null,
        listPeers: async ({ cfg, query, limit }) => {
            const a365Cfg = cfg.channels?.a365;
            const q = query?.trim().toLowerCase() || "";
            const ids = new Set();
            for (const entry of a365Cfg?.allowFrom ?? []) {
                const trimmed = String(entry).trim();
                if (trimmed && trimmed !== "*") {
                    ids.add(trimmed);
                }
            }
            return Array.from(ids)
                .filter((id) => (q ? id.toLowerCase().includes(q) : true))
                .slice(0, limit && limit > 0 ? limit : undefined)
                .map((id) => ({ kind: "user", id }));
        },
        listGroups: async () => [],
    },
    // Register Graph API tools for agent use
    agentTools: ({ cfg }) => {
        const a365Cfg = cfg?.channels?.a365;
        // Only provide tools if Graph API is configured (T1/T2 flow or callback)
        if (!isGraphConfigured(a365Cfg)) {
            return [];
        }
        return createGraphTools(a365Cfg);
    },
    outbound: a365Outbound,
    status: {
        defaultRuntime: {
            accountId: DEFAULT_ACCOUNT_ID,
            running: false,
            lastStartAt: null,
            lastStopAt: null,
            lastError: null,
            port: null,
        },
        buildChannelSummary: ({ snapshot }) => ({
            configured: snapshot.configured ?? false,
            running: snapshot.running ?? false,
            lastStartAt: snapshot.lastStartAt ?? null,
            lastStopAt: snapshot.lastStopAt ?? null,
            lastError: snapshot.lastError ?? null,
            port: snapshot.port ?? null,
            probe: snapshot.probe,
            lastProbeAt: snapshot.lastProbeAt ?? null,
        }),
        probeAccount: async ({ cfg, account }) => {
            const a365Cfg = cfg.channels?.a365;
            const acctCfg = resolveAccountA365Config(a365Cfg, account.accountId);
            return probeA365(acctCfg);
        },
        buildAccountSnapshot: ({ account, runtime, probe }) => ({
            accountId: account.accountId,
            enabled: account.enabled,
            configured: account.configured,
            running: runtime?.running ?? false,
            lastStartAt: runtime?.lastStartAt ?? null,
            lastStopAt: runtime?.lastStopAt ?? null,
            lastError: runtime?.lastError ?? null,
            port: runtime?.port ?? null,
            probe,
        }),
    },
    gateway: {
        startAccount: async (ctx) => {
            const { monitorA365Provider } = await import("./monitor.js");
            const a365Cfg = ctx.cfg.channels?.a365;
            const port = a365Cfg?.webhook?.port ?? 3978;
            ctx.setStatus({ accountId: ctx.accountId, port });
            ctx.log?.info(`starting a365 provider (port ${port})`);
            return monitorA365Provider({
                cfg: ctx.cfg,
                runtime: ctx.runtime,
                abortSignal: ctx.abortSignal,
            });
        },
    },
};
