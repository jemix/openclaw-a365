import { DEFAULT_ACCOUNT_ID } from "openclaw/plugin-sdk/core";

export function normalizeA365AccountId(accountId?: string | null): string {
  const trimmed = accountId?.trim();
  return trimmed || DEFAULT_ACCOUNT_ID;
}

export function buildA365NamespacedPeerId(accountId: string | undefined, peerId: string): string {
  const normalizedPeerId = peerId.trim();
  if (!normalizedPeerId) return normalizeA365AccountId(accountId);
  return `${normalizeA365AccountId(accountId)}:${normalizedPeerId}`;
}

export function buildA365LookupKeys(peerId: string, accountId?: string): string[] {
  const normalizedPeerId = peerId.trim();
  if (!normalizedPeerId) return [];

  const normalizedAccountId = accountId ? normalizeA365AccountId(accountId) : undefined;
  const keys = new Set<string>();
  keys.add(normalizedPeerId);
  if (normalizedAccountId) {
    keys.add(buildA365NamespacedPeerId(normalizedAccountId, normalizedPeerId));
  }
  return [...keys];
}
