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
import { buildA365NamespacedPeerId, normalizeA365AccountId } from "./account-scope.js";

/**
 * SDK ConversationReference shape (from @microsoft/agents-activity).
 * We use a loose type to avoid tight coupling to the SDK version.
 */
export type StoredConversationReference = {
  activityId?: string;
  user?: Record<string, unknown>;
  agent?: Record<string, unknown>;
  conversation: { id: string; [key: string]: unknown };
  channelId: string;
  serviceUrl?: string;
  locale?: string;
};

type ConversationStoreEntry = {
  ref: StoredConversationReference;
  updatedAt: number;
  userAadId?: string;
  accountId?: string;
  peerId?: string;
};

type ConversationStore = {
  references: Record<string, ConversationStoreEntry>;
};

const STORE_DIR = join(homedir(), ".openclaw");
const STORE_PATH = join(STORE_DIR, "a365-conversations.json");

let storeCache: ConversationStore | null = null;

function resolveStoreKey(targetId: string, accountId?: string): string {
  return accountId ? buildA365NamespacedPeerId(accountId, targetId) : targetId;
}

function resolveEntry(
  store: ConversationStore,
  targetId: string,
  accountId?: string,
): ConversationStoreEntry | undefined {
  const scopedKey = resolveStoreKey(targetId, accountId);
  return store.references[scopedKey] ?? store.references[targetId];
}

async function loadStore(): Promise<ConversationStore> {
  if (storeCache) return storeCache;
  try {
    const data = JSON.parse(await readFile(STORE_PATH, "utf-8")) as ConversationStore;
    storeCache = data;
    return data;
  } catch {
    storeCache = { references: {} };
    return storeCache;
  }
}

async function saveStore(store: ConversationStore): Promise<void> {
  storeCache = store;
  await mkdir(STORE_DIR, { recursive: true });
  await writeFile(STORE_PATH, JSON.stringify(store, null, 2), "utf-8");
}

/**
 * Save a conversation reference for proactive messaging.
 */
export async function saveConversationReference(
  ref: StoredConversationReference,
  options?: {
    userAadId?: string;
    accountId?: string;
    peerId?: string;
  },
): Promise<void> {
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
export async function getConversationReference(
  conversationId: string,
  accountId?: string,
): Promise<StoredConversationReference | undefined> {
  const store = await loadStore();
  return resolveEntry(store, conversationId, accountId)?.ref;
}

/**
 * Get a stored conversation reference by user AAD ID.
 */
export async function getConversationReferenceByUser(
  userAadId: string,
  accountId?: string,
): Promise<StoredConversationReference | undefined> {
  const store = await loadStore();
  let best: { ref: StoredConversationReference; updatedAt: number } | undefined;
  const normalizedAccountId = accountId ? normalizeA365AccountId(accountId) : undefined;
  for (const entry of Object.values(store.references)) {
    if (normalizedAccountId && entry.accountId !== normalizedAccountId) continue;
    if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
      best = entry;
    }
  }
  return best?.ref;
}

/**
 * Get the accountId associated with a conversation.
 */
export async function getAccountIdForConversation(
  conversationId: string,
  accountId?: string,
): Promise<string | undefined> {
  const store = await loadStore();
  return resolveEntry(store, conversationId, accountId)?.accountId;
}

/**
 * Get the full store entry for a conversation (includes accountId, userAadId, etc.).
 */
export async function getConversationEntry(
  conversationId: string,
  accountId?: string,
): Promise<ConversationStoreEntry | undefined> {
  const store = await loadStore();
  return resolveEntry(store, conversationId, accountId);
}

/**
 * Get the full store entry by user AAD ID (most recent conversation).
 */
export async function getConversationEntryByUser(
  userAadId: string,
  accountId?: string,
): Promise<ConversationStoreEntry | undefined> {
  const store = await loadStore();
  let best: ConversationStoreEntry | undefined;
  const normalizedAccountId = accountId ? normalizeA365AccountId(accountId) : undefined;
  for (const entry of Object.values(store.references)) {
    if (normalizedAccountId && entry.accountId !== normalizedAccountId) continue;
    if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
      best = entry;
    }
  }
  return best;
}

/**
 * List all stored conversation references.
 */
export async function listConversationReferences(): Promise<
  Array<{ conversationId: string; updatedAt: number; userAadId?: string; accountId?: string }>
> {
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
export async function deleteConversationReference(conversationId: string): Promise<void> {
  const store = await loadStore();
  delete store.references[conversationId];
  await saveStore(store);
}

/**
 * Clear all stored conversation references.
 */
export async function clearConversationReferences(): Promise<void> {
  await saveStore({ references: {} });
}
