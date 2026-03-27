/**
 * Persistent conversation reference store for proactive messaging.
 *
 * Stores SDK ConversationReference objects (from activity.getConversationReference())
 * to a JSON file so they survive restarts. These references contain the agentic
 * identity metadata the SDK needs to acquire AU tokens for proactive sends.
 */
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";
import { buildA365LookupKeys, buildA365NamespacedPeerId, normalizeA365AccountId } from "./account-scope.js";
const STORE_DIR = join(homedir(), ".openclaw");
const STORE_PATH = join(STORE_DIR, "a365-conversations.json");
let storeCache = null;
function resolveStoreKey(targetId, accountId) {
    return accountId ? buildA365NamespacedPeerId(accountId, targetId) : targetId;
}
function resolveEntry(store, targetId, accountId) {
    for (const key of buildA365LookupKeys(targetId, accountId)) {
        const entry = store.references[key];
        if (entry) {
            return entry;
        }
    }
    return undefined;
}
async function loadStore() {
    if (storeCache)
        return storeCache;
    try {
        const data = JSON.parse(await readFile(STORE_PATH, "utf-8"));
        storeCache = data;
        return data;
    }
    catch {
        storeCache = { references: {} };
        return storeCache;
    }
}
async function saveStore(store) {
    storeCache = store;
    await mkdir(STORE_DIR, { recursive: true });
    await writeFile(STORE_PATH, JSON.stringify(store, null, 2), "utf-8");
}
/**
 * Save a conversation reference for proactive messaging.
 */
export async function saveConversationReference(ref, options) {
    const store = await loadStore();
    const normalizedAccountId = options?.accountId ? normalizeA365AccountId(options.accountId) : undefined;
    const peerId = options?.peerId?.trim() || ref.conversation.id;
    const key = resolveStoreKey(peerId, normalizedAccountId);
    store.references[key] = {
        ref,
        updatedAt: Date.now(),
        userAadId: options?.userAadId,
        accountId: normalizedAccountId,
        peerId,
    };
    await saveStore(store);
}
/**
 * Get a stored conversation reference by conversation ID.
 */
export async function getConversationReference(conversationId, accountId) {
    const store = await loadStore();
    return resolveEntry(store, conversationId, accountId)?.ref;
}
/**
 * Get a stored conversation reference by user AAD ID.
 */
export async function getConversationReferenceByUser(userAadId, accountId) {
    const store = await loadStore();
    let best;
    const normalizedAccountId = accountId ? normalizeA365AccountId(accountId) : undefined;
    for (const entry of Object.values(store.references)) {
        if (normalizedAccountId && entry.accountId !== normalizedAccountId)
            continue;
        if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
            best = entry;
        }
    }
    return best?.ref;
}
/**
 * Get the accountId associated with a conversation.
 */
export async function getAccountIdForConversation(conversationId, accountId) {
    const store = await loadStore();
    return resolveEntry(store, conversationId, accountId)?.accountId;
}
/**
 * Get the full store entry for a conversation (includes accountId, userAadId, etc.).
 */
export async function getConversationEntry(conversationId, accountId) {
    const store = await loadStore();
    return resolveEntry(store, conversationId, accountId);
}
/**
 * Get the full store entry by user AAD ID (most recent conversation).
 */
export async function getConversationEntryByUser(userAadId, accountId) {
    const store = await loadStore();
    let best;
    const normalizedAccountId = accountId ? normalizeA365AccountId(accountId) : undefined;
    for (const entry of Object.values(store.references)) {
        if (normalizedAccountId && entry.accountId !== normalizedAccountId)
            continue;
        if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
            best = entry;
        }
    }
    return best;
}
/**
 * List all stored conversation references.
 */
export async function listConversationReferences() {
    const store = await loadStore();
    return Object.entries(store.references).map(([id, entry]) => ({
        conversationId: id,
        updatedAt: entry.updatedAt,
        userAadId: entry.userAadId,
        accountId: entry.accountId,
        peerId: entry.peerId,
    }));
}
/**
 * Delete a stored conversation reference.
 */
export async function deleteConversationReference(conversationId, accountId) {
    const store = await loadStore();
    for (const key of buildA365LookupKeys(conversationId, accountId)) {
        delete store.references[key];
    }
    await saveStore(store);
}
/**
 * Clear all stored conversation references.
 */
export async function clearConversationReferences() {
    await saveStore({ references: {} });
}
