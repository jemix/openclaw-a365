import {
  createChatChannelPlugin,
  createChannelPluginBase,
  DEFAULT_ACCOUNT_ID,
} from "openclaw/plugin-sdk/core";
import type { OpenClawConfig } from "openclaw/plugin-sdk/core";
import type { A365Config, A365AccountConfig, A365Probe, ResolvedA365Account } from "./types.js";
import { resolveA365Credentials, resolveGraphTokenConfig, resolveTokenCallbackConfig, getGraphToken } from "./token.js";
import { createGraphTools } from "./graph-tools.js";
import { a365Outbound, normalizeA365MessagingTarget } from "./outbound.js";

/**
 * Resolve the effective A365Config for a specific account.
 *
 * Multi-account mode: merges top-level defaults with per-account overrides.
 * Single-account mode (no `accounts` key): returns the flat config as-is.
 */
export function resolveAccountA365Config(
  a365Cfg: A365Config | undefined,
  accountId?: string | null,
): A365Config | undefined {
  if (!a365Cfg) return undefined;
  const accounts = a365Cfg.accounts;
  if (!accounts || !accountId || accountId === DEFAULT_ACCOUNT_ID && !accounts[DEFAULT_ACCOUNT_ID]) {
    return a365Cfg;
  }
  const acct = accounts[accountId];
  if (!acct) return undefined;
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

function listA365AccountIds(a365Cfg?: A365Config): string[] {
  if (a365Cfg?.accounts && Object.keys(a365Cfg.accounts).length > 0) {
    return Object.keys(a365Cfg.accounts);
  }
  return [DEFAULT_ACCOUNT_ID];
}

function isGraphConfigured(cfg?: A365Config): boolean {
  return Boolean(resolveGraphTokenConfig(cfg) || resolveTokenCallbackConfig(cfg));
}

async function probeA365(cfg?: A365Config): Promise<A365Probe> {
  const creds = resolveA365Credentials(cfg);
  if (!creds) {
    return { ok: false, error: "Bot Framework credentials not configured" };
  }
  const graphConfigured = isGraphConfigured(cfg);
  let graphConnected = false;
  if (graphConfigured && cfg?.agentIdentity) {
    try {
      const token = await getGraphToken(cfg, cfg.agentIdentity);
      graphConnected = Boolean(token);
    } catch {
      graphConnected = false;
    }
  }
  return { ok: true, botId: creds.appId, graphConnected, owner: cfg?.owner };
}

// ── Account resolution for the SDK ──

function resolveAccount(cfg: OpenClawConfig, accountId?: string | null): ResolvedA365Account {
  const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
  const effectiveId = accountId || DEFAULT_ACCOUNT_ID;
  const acctCfg = resolveAccountA365Config(a365Cfg, effectiveId);
  return {
    accountId: effectiveId,
    enabled: acctCfg?.enabled !== false,
    configured: Boolean(resolveA365Credentials(acctCfg)),
    owner: acctCfg?.owner,
  };
}

// ── Build the channel plugin ──

export const a365Plugin = createChatChannelPlugin<ResolvedA365Account, A365Probe>({
  base: createChannelPluginBase<ResolvedA365Account>({
    id: "a365",
    meta: {
      label: "Microsoft 365 Agents",
      selectionLabel: "Microsoft 365 Agents (A365)",
      docsPath: "/channels/a365",
      docsLabel: "a365",
      blurb: "Native A365 channel with Graph API tools for calendar and email.",
      aliases: ["m365agents", "agents365"] as string[],
      order: 55,
    },
    capabilities: {
      chatTypes: ["direct", "channel", "thread"],
      threads: true,
      media: true,
    },
    setup: {
      resolveAccountId: ({ cfg, accountId }) => {
        if (accountId) return accountId;
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        const ids = listA365AccountIds(a365Cfg);
        return ids[0] ?? DEFAULT_ACCOUNT_ID;
      },
      applyAccountConfig: ({ cfg, accountId }) => {
        const a365Cfg = (cfg.channels?.a365 ?? {}) as A365Config;
        const accounts = a365Cfg.accounts;
        if (accounts && accountId && accountId !== DEFAULT_ACCOUNT_ID) {
          return {
            ...cfg,
            channels: {
              ...cfg.channels,
              a365: { ...a365Cfg, accounts: { ...accounts, [accountId]: { ...accounts[accountId], enabled: true } } },
            },
          };
        }
        return { ...cfg, channels: { ...cfg.channels, a365: { ...a365Cfg, enabled: true } } };
      },
    },
    agentPrompt: {
      messageToolHints: ({ cfg }) => {
        const a365Cfg = cfg?.channels?.a365 as A365Config | undefined;
        const timezone = a365Cfg?.businessHours?.timezone || "America/Los_Angeles";
        const now = new Date();
        const formatter = new Intl.DateTimeFormat("en-US", {
          timeZone: timezone,
          year: "numeric", month: "2-digit", day: "2-digit",
          hour: "2-digit", minute: "2-digit", hour12: false,
        });
        const currentDateTime = formatter.format(now);
        const dateOnly = new Date().toLocaleDateString("en-CA", { timeZone: timezone });

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
    reload: { configPrefixes: ["channels.a365"] },
    config: {
      listAccountIds: (cfg) => {
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        return listA365AccountIds(a365Cfg);
      },
      resolveAccount: (cfg, accountId) => resolveAccount(cfg, accountId),
      defaultAccountId: (cfg) => {
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        const ids = listA365AccountIds(a365Cfg);
        return ids[0] ?? DEFAULT_ACCOUNT_ID;
      },
      setAccountEnabled: ({ cfg, accountId, enabled }) => {
        const a365Cfg = (cfg.channels?.a365 ?? {}) as A365Config;
        const accounts = a365Cfg.accounts;
        if (accounts && accountId && accountId !== DEFAULT_ACCOUNT_ID) {
          return {
            ...cfg,
            channels: {
              ...cfg.channels,
              a365: {
                ...a365Cfg,
                accounts: { ...accounts, [accountId]: { ...accounts[accountId], enabled } },
              },
            },
          };
        }
        return { ...cfg, channels: { ...cfg.channels, a365: { ...a365Cfg, enabled } } };
      },
      deleteAccount: ({ cfg, accountId }) => {
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        const accounts = a365Cfg?.accounts;
        if (accounts && accountId && accounts[accountId]) {
          const nextAccounts = { ...accounts };
          delete nextAccounts[accountId];
          const nextA365: A365Config = { ...a365Cfg };
          if (Object.keys(nextAccounts).length > 0) {
            nextA365.accounts = nextAccounts;
          } else {
            delete nextA365.accounts;
          }
          return { ...cfg, channels: { ...cfg.channels, a365: nextA365 } };
        }
        const next = { ...cfg } as OpenClawConfig;
        const nextChannels = { ...cfg.channels };
        delete nextChannels.a365;
        next.channels = Object.keys(nextChannels).length > 0 ? nextChannels : undefined;
        return next;
      },
      isConfigured: (_account, cfg) => {
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        const acctCfg = resolveAccountA365Config(a365Cfg, _account.accountId);
        return Boolean(resolveA365Credentials(acctCfg));
      },
      describeAccount: (account) => ({
        accountId: account.accountId,
        enabled: account.enabled,
        configured: account.configured,
      }),
      resolveAllowFrom: ({ cfg, accountId }) => {
        const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
        const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
        return acctCfg?.allowFrom?.map(String) ?? [];
      },
      formatAllowFrom: ({ allowFrom }) =>
        allowFrom.map((entry) => String(entry).trim()).filter(Boolean).map((entry) => entry.toLowerCase()),
    },
  }),

  // ── DM Security ──
  security: {
    dm: {
      channelKey: "a365",
      resolvePolicy: (account) => {
        // We need the full config for policy resolution; use a lightweight approach
        return undefined; // Falls back to defaultPolicy
      },
      resolveAllowFrom: (account) => [],
      defaultPolicy: "pairing",
    },
    collectWarnings: ({ cfg, accountId }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
      const groupPolicy = acctCfg?.groupPolicy ?? "allowlist";
      if (groupPolicy !== "open") return [];
      return [
        `- A365 groups: groupPolicy="open" allows any member to trigger. Set channels.a365.groupPolicy="allowlist" + channels.a365.groupAllowFrom to restrict senders.`,
      ];
    },
  },

  // ── Pairing ──
  pairing: {
    text: {
      idLabel: "a365UserId",
      message: "Send this code to verify your identity:",
      normalizeAllowEntry: (entry) => entry.replace(/^(a365|user):/i, ""),
      notify: async () => {
        // A365 pairing notification is handled via Bot Framework — no-op here
      },
    },
  },

  // ── Threading ──
  threading: {
    topLevelReplyToMode: "reply",
    buildToolContext: ({ context, hasRepliedRef }) => ({
      currentChannelId: context.To?.trim() || undefined,
      currentThreadTs: context.ReplyToId,
      hasRepliedRef,
    }),
  },

  // ── Outbound ──
  outbound: a365Outbound,
});

// ── Attach adapters that don't fit the builder pattern ──

// These are set directly because createChatChannelPlugin doesn't have
// explicit slots for them — they pass through from the base object.

// Messaging adapter
(a365Plugin as any).messaging = {
  normalizeTarget: normalizeA365MessagingTarget,
  targetResolver: {
    looksLikeId: (raw: string) => {
      const trimmed = raw.trim();
      if (!trimmed) return false;
      if (/^conversation:/i.test(trimmed)) return true;
      if (/^user:/i.test(trimmed)) {
        const id = trimmed.slice("user:".length).trim();
        return /^[0-9a-fA-F-]{16,}$/.test(id);
      }
      return trimmed.includes("@thread") || trimmed.includes(":");
    },
    hint: "<conversationId|user:ID|conversation:ID>",
  },
};

// Directory adapter
(a365Plugin as any).directory = {
  self: async () => null,
  listPeers: async ({ cfg, query, limit }: any) => {
    const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
    const q = query?.trim().toLowerCase() || "";
    const ids = new Set<string>();
    for (const entry of a365Cfg?.allowFrom ?? []) {
      const trimmed = String(entry).trim();
      if (trimmed && trimmed !== "*") ids.add(trimmed);
    }
    return Array.from(ids)
      .filter((id) => (q ? id.toLowerCase().includes(q) : true))
      .slice(0, limit && limit > 0 ? limit : undefined)
      .map((id) => ({ kind: "user", id }) as const);
  },
  listGroups: async () => [],
};

// Agent tools (Graph API)
(a365Plugin as any).agentTools = ({ cfg }: any) => {
  const a365Cfg = cfg?.channels?.a365 as A365Config | undefined;
  if (!isGraphConfigured(a365Cfg)) return [];
  return createGraphTools(a365Cfg);
};

// Status adapter
(a365Plugin as any).status = {
  defaultRuntime: {
    accountId: DEFAULT_ACCOUNT_ID,
    running: false,
    lastStartAt: null,
    lastStopAt: null,
    lastError: null,
    port: null,
  },
  buildChannelSummary: ({ snapshot }: any) => ({
    configured: snapshot.configured ?? false,
    running: snapshot.running ?? false,
    lastStartAt: snapshot.lastStartAt ?? null,
    lastStopAt: snapshot.lastStopAt ?? null,
    lastError: snapshot.lastError ?? null,
    port: snapshot.port ?? null,
    probe: snapshot.probe,
    lastProbeAt: snapshot.lastProbeAt ?? null,
  }),
  probeAccount: async ({ cfg, account }: any) => {
    const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
    const acctCfg = resolveAccountA365Config(a365Cfg, account.accountId);
    return probeA365(acctCfg);
  },
  buildAccountSnapshot: ({ account, runtime, probe }: any) => ({
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
};

// Gateway adapter
(a365Plugin as any).gateway = {
  startAccount: async (ctx: any) => {
    const { monitorA365Provider } = await import("./monitor.js");
    const a365Cfg = ctx.cfg.channels?.a365 as A365Config | undefined;
    const port = a365Cfg?.webhook?.port ?? 3978;
    ctx.setStatus({ accountId: ctx.accountId, port });
    ctx.log?.info(`starting a365 provider (port ${port})`);
    return monitorA365Provider({
      cfg: ctx.cfg,
      runtime: ctx.runtime,
      abortSignal: ctx.abortSignal,
    });
  },
};
