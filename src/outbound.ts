type ChannelOutboundAdapter = any;
import type { A365Config, A365MessageMetadata } from "./types.js";
import { getA365Runtime } from "./runtime.js";
import {
  getAdapter,
  getBlueprintClientId,
  getAdapterForAccount,
  getAdapterByRecipientId,
} from "./adapter-store.js";
import {
  getConversationReference,
  getConversationReferenceByUser,
  getAccountIdForConversation,
  getConversationEntryByUser,
  type StoredConversationReference,
} from "./conversation-store.js";

type ResolvedReference = {
  ref: StoredConversationReference;
  accountId?: string;
};

/**
 * Resolve a stored conversation reference from the provided params.
 * Tries conversationId lookup first, then falls back to userAadId lookup.
 * Returns the reference and its associated accountId (if stored).
 */
async function resolveStoredReference(
  to: string | undefined,
  metadata: A365MessageMetadata | undefined,
): Promise<ResolvedReference | undefined> {
  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  const conversationId = to || metadata?.conversationId;
  if (!conversationId) return undefined;

  log.info(`Looking up stored conversation reference for: ${conversationId}`);
  let ref = await getConversationReference(conversationId);
  let accountId = ref ? await getAccountIdForConversation(conversationId) : undefined;

  // If direct lookup fails and it looks like a userAadId (UUID without colons/@),
  // try looking up by userAadId as a fallback
  if (!ref && !conversationId.includes(":") && !conversationId.includes("@")) {
    log.info(`Direct lookup failed, trying userAadId lookup for ${conversationId}`);
    ref = await getConversationReferenceByUser(conversationId);
    if (ref) {
      const entry = await getConversationEntryByUser(conversationId);
      accountId = entry?.accountId;
    }
  }

  if (ref) {
    log.info(`Found stored reference: conversationId=${ref.conversation?.id}, accountId=${accountId ?? "default"}`);
    return { ref, accountId };
  }

  log.warn(`No stored reference found for: ${conversationId}`);
  return undefined;
}

/**
 * Send a proactive message using the SDK's adapter.continueConversation().
 *
 * Multi-account: resolves the correct adapter by accountId (from conversation-store)
 * or by ref.agent.id (recipientId). Falls back to the default adapter.
 *
 * Per the SDK author, two things are required:
 * 1. The ConversationReference must come from an AU-based (agentic user) inbound request
 *    (captured via activity.getConversationReference() in monitor.ts)
 * 2. The clientId must be the Blueprint Client App ID (not the bot's own app ID)
 *
 * With these in place, the SDK handles T1/T2/AU token acquisition internally.
 */
async function sendViaAdapter(params: {
  ref: StoredConversationReference;
  accountId?: string;
  sendFn: (context: { sendActivity: (activity: unknown) => Promise<{ id?: string }> }) => Promise<{ id?: string }>;
}): Promise<{ id?: string }> {
  const { ref, accountId, sendFn } = params;
  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  // Try to resolve the correct adapter:
  // 1. By accountId from conversation-store
  // 2. By ref.agent.id (recipientId) from adapter-store
  // 3. Fall back to default adapter
  let adapter: import("@microsoft/agents-hosting").CloudAdapter | null = null;
  let blueprintClientId: string | null = null;

  if (accountId) {
    const entry = getAdapterForAccount(accountId);
    if (entry) {
      adapter = entry.adapter;
      blueprintClientId = entry.blueprintClientId;
      log.debug("resolved adapter by accountId", { accountId });
    }
  }

  if (!adapter && ref.agent) {
    const agentId = (ref.agent as Record<string, unknown>).id as string | undefined;
    if (agentId) {
      const entry = getAdapterByRecipientId(agentId);
      if (entry) {
        adapter = entry.adapter;
        blueprintClientId = entry.blueprintClientId;
        log.debug("resolved adapter by agent.id", { agentId });
      }
    }
  }

  if (!adapter) {
    adapter = getAdapter();
    blueprintClientId = getBlueprintClientId();
    log.debug("using default adapter");
  }

  if (!adapter) {
    throw new Error("CloudAdapter not initialized. Bot must receive at least one message first.");
  }

  if (!blueprintClientId) {
    throw new Error("Blueprint Client App ID not configured.");
  }

  log.info("Sending via adapter.continueConversation", {
    blueprintClientId,
    conversationId: ref.conversation?.id,
    agentRole: (ref.agent as Record<string, unknown>)?.role,
    accountId: accountId ?? "default",
  });

  let result: { id?: string } = {};

  await adapter.continueConversation(blueprintClientId, ref, async (context) => {
    result = await sendFn(context);
  });

  return result;
}

/**
 * Send a message to a conversation using the SDK's proactive messaging.
 */
export async function sendMessageA365(params: {
  cfg: unknown;
  to: string;
  text: string;
  serviceUrl?: string;
  metadata?: A365MessageMetadata;
}): Promise<{ ok: boolean; messageId?: string; conversationId?: string; error?: string }> {
  const { to, text, metadata } = params;
  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  log.info(`sendMessageA365 called: to=${to} hasMetadata=${!!metadata}`);

  const resolved = await resolveStoredReference(to, metadata);
  if (!resolved) {
    return { ok: false, error: "No stored conversation reference. User must message the bot first." };
  }

  try {
    const result = await sendViaAdapter({
      ref: resolved.ref,
      accountId: resolved.accountId,
      sendFn: async (context) => {
        const res = await context.sendActivity({ type: "message", text });
        return { id: res?.id };
      },
    });

    log.info("Proactive message sent successfully", { messageId: result.id });
    return {
      ok: true,
      messageId: result.id,
      conversationId: resolved.ref.conversation?.id,
    };
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    log.error(`Proactive message failed: ${detail}`);
    return { ok: false, error: detail };
  }
}

/**
 * Send an Adaptive Card to a conversation using the SDK's proactive messaging.
 */
export async function sendAdaptiveCardA365(params: {
  cfg: unknown;
  to: string;
  card: Record<string, unknown>;
  serviceUrl?: string;
  metadata?: A365MessageMetadata;
}): Promise<{ ok: boolean; messageId?: string; conversationId?: string; error?: string }> {
  const { to, card, metadata } = params;
  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  const resolved = await resolveStoredReference(to, metadata);
  if (!resolved) {
    return { ok: false, error: "No stored conversation reference. User must message the bot first." };
  }

  try {
    const result = await sendViaAdapter({
      ref: resolved.ref,
      accountId: resolved.accountId,
      sendFn: async (context) => {
        const res = await context.sendActivity({
          type: "message",
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: card,
            },
          ],
        });
        return { id: res?.id };
      },
    });

    log.info("Card sent successfully", { messageId: result.id });
    return {
      ok: true,
      messageId: result.id,
      conversationId: resolved.ref.conversation?.id,
    };
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    log.error(`Card send failed: ${detail}`);
    return { ok: false, error: detail };
  }
}

/**
 * Normalize an A365 target to a conversation ID.
 * Handles various formats that might come from session storage.
 */
function normalizeA365Target(to: string | undefined): string | undefined {
  if (!to) return undefined;
  const trimmed = to.trim();
  if (!trimmed) return undefined;

  const prefixes = ["user:", "conversation:", "a365:", "a365:group:"];
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      return trimmed.slice(prefix.length);
    }
  }

  return trimmed;
}

/**
 * A365 outbound adapter for sending messages.
 */
export const a365Outbound: ChannelOutboundAdapter = {
  deliveryMode: "direct",
  textChunkLimit: 4000,

  resolveTarget: ({ to, allowFrom }) => {
    const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
    log.info("resolveTarget called", { to, allowFromCount: allowFrom?.length ?? 0 });

    const normalized = normalizeA365Target(to);

    if (normalized) {
      log.info("resolveTarget success", { normalized });
      return { ok: true, to: normalized };
    }

    const allowList = (allowFrom ?? []).map((entry) => String(entry).trim()).filter(Boolean);
    if (allowList.length > 0) {
      const fallback = normalizeA365Target(allowList[0]);
      if (fallback) {
        log.info("resolveTarget fallback", { fallback });
        return { ok: true, to: fallback };
      }
    }

    log.warn("resolveTarget failed - no target");
    return {
      ok: false,
      error: new Error("No A365 conversation target specified. User must message the bot first."),
    };
  },

  sendText: async ({ cfg, to, text }) => {
    const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
    log.info("sendText called", { to, textLength: text?.length ?? 0 });

    const result = await sendMessageA365({ cfg, to, text });
    return {
      channel: "a365" as const,
      messageId: result.messageId ?? "",
      conversationId: result.conversationId,
    };
  },

  sendMedia: async ({ cfg, to, text, mediaUrl }) => {
    const messageText = mediaUrl ? `${text}\n\n${mediaUrl}` : text;
    const result = await sendMessageA365({ cfg, to, text: messageText });
    return {
      channel: "a365" as const,
      messageId: result.messageId ?? "",
      conversationId: result.conversationId,
    };
  },
};

/**
 * Normalize A365 messaging target.
 */
export function normalizeA365MessagingTarget(raw: string): string | undefined {
  const trimmed = raw.trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed.toLowerCase().startsWith("conversation:")) {
    return trimmed.slice("conversation:".length).trim() || undefined;
  }

  if (trimmed.toLowerCase().startsWith("user:")) {
    return `user:${trimmed.slice("user:".length).trim()}`;
  }

  if (trimmed.includes("@") || trimmed.includes(":")) {
    return trimmed;
  }

  return trimmed;
}
