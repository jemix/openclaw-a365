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
const STORE_DIR = join(homedir(), ".openclaw");
const STORE_PATH = join(STORE_DIR, "a365-conversations.json");
let storeCache = null;
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
export async function saveConversationReference(ref, userAadId, accountId) {
    const store = await loadStore();
    const key = ref.conversation.id;
    store.references[key] = { ref, updatedAt: Date.now(), userAadId, accountId };
    await saveStore(store);
}
/**
 * Get a stored conversation reference by conversation ID.
 */
export async function getConversationReference(conversationId) {
    const store = await loadStore();
    return store.references[conversationId]?.ref;
}
/**
 * Get a stored conversation reference by user AAD ID.
 */
export async function getConversationReferenceByUser(userAadId) {
    const store = await loadStore();
    let best;
    for (const entry of Object.values(store.references)) {
        if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
            best = entry;
        }
    }
    return best?.ref;
}
/**
 * Get the accountId associated with a conversation.
 */
export async function getAccountIdForConversation(conversationId) {
    const store = await loadStore();
    return store.references[conversationId]?.accountId;
}
/**
 * Get the full store entry for a conversation (includes accountId, userAadId, etc.).
 */
export async function getConversationEntry(conversationId) {
    const store = await loadStore();
    return store.references[conversationId];
}
/**
 * Get the full store entry by user AAD ID (most recent conversation).
 */
export async function getConversationEntryByUser(userAadId) {
    const store = await loadStore();
    let best;
    for (const entry of Object.values(store.references)) {
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
    }));
}
/**
 * Delete a stored conversation reference.
 */
export async function deleteConversationReference(conversationId) {
    const store = await loadStore();
    delete store.references[conversationId];
    await saveStore(store);
}
/**
 * Clear all stored conversation references.
 */
export async function clearConversationReferences() {
    await saveStore({ references: {} });
}
