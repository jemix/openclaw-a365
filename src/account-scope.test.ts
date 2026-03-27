import { describe, expect, it } from "vitest";
import { buildA365LookupKeys, buildA365NamespacedPeerId, normalizeA365AccountId } from "./account-scope.js";

describe("account-scope", () => {
  it("defaults empty account ids to default", () => {
    expect(normalizeA365AccountId(undefined)).toBe("default");
    expect(normalizeA365AccountId("")).toBe("default");
  });

  it("builds namespaced peer ids", () => {
    expect(buildA365NamespacedPeerId("aila", "user-123")).toBe("aila:user-123");
    expect(buildA365NamespacedPeerId(undefined, "conv-456")).toBe("default:conv-456");
  });

  it("provides scoped and legacy lookup keys", () => {
    expect(buildA365LookupKeys("abc", "aila")).toEqual(["abc", "aila:abc"]);
  });
});
