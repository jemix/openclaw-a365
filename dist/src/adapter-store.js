/** Primary index: accountId → AdapterEntry */
const adapterMap = new Map();
/** Secondary index: appId → accountId */
const appIdIndex = new Map();
/** Secondary index: recipientId → accountId (for routing inbound by activity.recipient.id) */
const recipientIdIndex = new Map();
/** Track insertion order for default fallback */
let firstAccountId = null;
export function registerAdapter(params) {
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
export function registerRecipientId(recipientId, accountId) {
    recipientIdIndex.set(recipientId, accountId);
}
/** Get adapter entry by accountId. */
export function getAdapterForAccount(accountId) {
    return adapterMap.get(accountId) ?? null;
}
/** Get adapter entry by recipient.id (for routing inbound activities). */
export function getAdapterByRecipientId(recipientId) {
    const accountId = recipientIdIndex.get(recipientId);
    if (!accountId)
        return null;
    return adapterMap.get(accountId) ?? null;
}
/** Resolve accountId from a recipientId. */
export function resolveAccountIdByRecipientId(recipientId) {
    return recipientIdIndex.get(recipientId);
}
/** Get all registered account IDs. */
export function getRegisteredAccountIds() {
    return Array.from(adapterMap.keys());
}
// --- Backward-compatible single-adapter API ---
/** Get the default (first registered) adapter. */
export function getAdapter() {
    if (!firstAccountId)
        return null;
    return adapterMap.get(firstAccountId)?.adapter ?? null;
}
/** Get the default (first registered) blueprint client ID. */
export function getBlueprintClientId() {
    if (!firstAccountId)
        return null;
    return adapterMap.get(firstAccountId)?.blueprintClientId ?? null;
}
// --- Legacy setters (kept for any code that still calls them) ---
export function setAdapter(adapter) {
    // If no adapters registered yet, create a default entry
    if (!firstAccountId) {
        firstAccountId = "__default__";
        adapterMap.set("__default__", { adapter, blueprintClientId: "", agentApp: null });
    }
    else {
        const entry = adapterMap.get(firstAccountId);
        if (entry) {
            entry.adapter = adapter;
        }
    }
}
export function setBlueprintClientId(clientId) {
    if (!firstAccountId)
        return;
    const entry = adapterMap.get(firstAccountId);
    if (entry) {
        entry.blueprintClientId = clientId;
    }
}
