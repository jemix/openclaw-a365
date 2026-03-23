/**
 * Stores CloudAdapters for multi-account proactive messaging.
 *
 * Supports multiple adapters keyed by accountId, with secondary indexes
 * by appId and recipientId for routing inbound activities and outbound sends.
 *
 * Backward compatible: getAdapter() / getBlueprintClientId() return the first/default entry.
 */
import type { CloudAdapter } from "@microsoft/agents-hosting";

type AdapterEntry = {
  adapter: CloudAdapter;
  blueprintClientId: string;
  agentApp: unknown;
};

/** Primary index: accountId → AdapterEntry */
const adapterMap = new Map<string, AdapterEntry>();

/** Secondary index: appId → accountId */
const appIdIndex = new Map<string, string>();

/** Secondary index: recipientId → accountId (for routing inbound by activity.recipient.id) */
const recipientIdIndex = new Map<string, string>();

/** Track insertion order for default fallback */
let firstAccountId: string | null = null;

export function registerAdapter(params: {
  accountId: string;
  appId: string;
  adapter: CloudAdapter;
  blueprintClientId: string;
  agentApp: unknown;
}): void {
  const { accountId, appId, adapter, blueprintClientId, agentApp } = params;
  adapterMap.set(accountId, { adapter, blueprintClientId, agentApp });
  appIdIndex.set(appId, accountId);
  // recipientId is typically the same as appId in Bot Framework
  recipientIdIndex.set(appId, accountId);
  if (!firstAccountId) {
    firstAccountId = accountId;
  }
}

/**
 * Register an additional recipientId → accountId mapping.
 * Useful when the bot's recipient.id differs from its appId.
 */
export function registerRecipientId(recipientId: string, accountId: string): void {
  recipientIdIndex.set(recipientId, accountId);
}

/** Get adapter entry by accountId. */
export function getAdapterForAccount(accountId: string): AdapterEntry | null {
  return adapterMap.get(accountId) ?? null;
}

/** Get adapter entry by recipient.id (for routing inbound activities). */
export function getAdapterByRecipientId(recipientId: string): AdapterEntry | null {
  const accountId = recipientIdIndex.get(recipientId);
  if (!accountId) return null;
  return adapterMap.get(accountId) ?? null;
}

/** Resolve accountId from a recipientId. */
export function resolveAccountIdByRecipientId(recipientId: string): string | undefined {
  return recipientIdIndex.get(recipientId);
}

/** Get all registered account IDs. */
export function getRegisteredAccountIds(): string[] {
  return Array.from(adapterMap.keys());
}

// --- Backward-compatible single-adapter API ---

/** Get the default (first registered) adapter. */
export function getAdapter(): CloudAdapter | null {
  if (!firstAccountId) return null;
  return adapterMap.get(firstAccountId)?.adapter ?? null;
}

/** Get the default (first registered) blueprint client ID. */
export function getBlueprintClientId(): string | null {
  if (!firstAccountId) return null;
  return adapterMap.get(firstAccountId)?.blueprintClientId ?? null;
}

// --- Legacy setters (kept for any code that still calls them) ---

export function setAdapter(adapter: CloudAdapter): void {
  // If no adapters registered yet, create a default entry
  if (!firstAccountId) {
    firstAccountId = "__default__";
    adapterMap.set("__default__", { adapter, blueprintClientId: "", agentApp: null });
  } else {
    const entry = adapterMap.get(firstAccountId);
    if (entry) {
      entry.adapter = adapter;
    }
  }
}

export function setBlueprintClientId(clientId: string): void {
  if (!firstAccountId) return;
  const entry = adapterMap.get(firstAccountId);
  if (entry) {
    entry.blueprintClientId = clientId;
  }
}
