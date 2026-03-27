import { DEFAULT_ACCOUNT_ID } from "openclaw/plugin-sdk/core";
export function normalizeA365AccountId(accountId) {
    const trimmed = accountId?.trim();
    return trimmed || DEFAULT_ACCOUNT_ID;
}
export function buildA365NamespacedPeerId(accountId, peerId) {
    const normalizedPeerId = peerId.trim();
    if (!normalizedPeerId)
        return normalizeA365AccountId(accountId);
    return `${normalizeA365AccountId(accountId)}:${normalizedPeerId}`;
}
