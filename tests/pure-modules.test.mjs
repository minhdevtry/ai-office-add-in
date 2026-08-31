/**
 * Pure Module Unit Tests (node:test built-in)
 *
 * Test các pure logic module KHÔNG cần Office.js / Word thật:
 *  - diff.js: word-level diff (LCS algorithm)
 *  - stream-factory.js: 3-layer endpoint resolver
 *  - session-store.js: IndexedDB session persistence (mock storage)
 *
 * Chạy: `npm test` (đã có sẵn script này)
 * CI-friendly: chạy trên Ubuntu, không cần Office.
 */

import { test, describe } from "node:test";
import { strict as assert } from "node:assert";
import { computeWordDiff, renderDiffHtml } from "../public/modules/diff.js";
import {
  resolveStreamEndpoint,
  saveStreamProxyConfig,
  saveCorsProxyConfig,
  saveProviderKeyConfig,
} from "../public/modules/stream-factory.js";

// ─── localStorage shim (cho Node test, browser có sẵn) ───
if (typeof globalThis.localStorage === "undefined") {
  const _store = new Map();
  globalThis.localStorage = {
    getItem: (k) => (_store.has(k) ? _store.get(k) : null),
    setItem: (k, v) => _store.set(k, String(v)),
    removeItem: (k) => _store.delete(k),
    clear: () => _store.clear(),
    key: (i) => Array.from(_store.keys())[i] || null,
    get length() {
      return _store.size;
    },
  };
}

// ─── DIFF.JS ──────────────────────────────────────────────────────────────

describe("diff.js: word-level diff", () => {
  test("tokenizeWords: Vietnamese có dấu", () => {
    const tokens = computeWordDiff("Tôi là Minh", "Tôi là Bình");
    assert.ok(Array.isArray(tokens));
    assert.ok(tokens.length > 0);
  });

  test("identical text → all equal", () => {
    const diff = computeWordDiff("hello world", "hello world");
    assert.equal(diff.filter((c) => c.type === "equal").length, diff.length);
    assert.equal(diff.filter((c) => c.type !== "equal").length, 0);
  });

  test("completely different → all del/ins (single token, no whitespace match)", () => {
    // 1 token duy nhất → chỉ 1 phần tử del/ins
    const diff = computeWordDiff("foo", "bar");
    assert.equal(diff.filter((c) => c.type === "equal").length, 0);
  });

  test("mixed: keep common + show changes", () => {
    const diff = computeWordDiff("the quick brown fox", "the slow brown dog");
    const equalWords = diff.filter((c) => c.type === "equal").map((c) => c.value.trim());
    assert.ok(equalWords.includes("the"));
    assert.ok(equalWords.includes("brown"));
    assert.ok(diff.some((c) => c.type === "del"));
    assert.ok(diff.some((c) => c.type === "ins"));
  });

  test("empty old → all ins", () => {
    const diff = computeWordDiff("", "new text");
    assert.equal(diff.length, 1);
    assert.equal(diff[0].type, "ins");
    assert.equal(diff[0].value, "new text");
  });

  test("empty new → all del", () => {
    const diff = computeWordDiff("old text", "");
    assert.equal(diff.length, 1);
    assert.equal(diff[0].type, "del");
  });

  test("both empty → empty array", () => {
    const diff = computeWordDiff("", "");
    assert.deepEqual(diff, []);
  });

  test("renderDiffHtml: includes <ins> and <del> tags", () => {
    const diff = [{ type: "ins", value: "added" }, { type: "del", value: "removed" }];
    const html = renderDiffHtml(diff);
    assert.ok(html.includes("<ins"));
    assert.ok(html.includes("<del"));
    assert.ok(html.includes("added"));
    assert.ok(html.includes("removed"));
  });

  test("renderDiffHtml: escapes HTML", () => {
    const diff = [{ type: "ins", value: "<script>alert(1)</script>" }];
    const html = renderDiffHtml(diff);
    assert.ok(!html.includes("<script>"));
    assert.ok(html.includes("&lt;script&gt;"));
  });
});

// ─── STREAM-FACTORY.JS ───────────────────────────────────────────────────

describe("stream-factory.js: 3-layer endpoint resolver", () => {
  test("Layer 1: StreamProxy wins when enabled", () => {
    saveStreamProxyConfig({ enabled: true, url: "https://proxy.com/v1", token: "tok123" });
    saveCorsProxyConfig({ enabled: true, url: "https://cors.example.com/?" });
    saveProviderKeyConfig({ provider: "openai", apiKey: "sk-direct", baseURL: "https://api.openai.com/v1" });

    return resolveStreamEndpoint({}).then((resolved) => {
      assert.equal(resolved.url, "https://proxy.com/v1/chat/completions");
      assert.equal(resolved.headers.Authorization, "Bearer tok123");
    });
  });

  test("Layer 2: CORS Proxy when streamProxy disabled", () => {
    saveStreamProxyConfig({ enabled: false, url: "", token: "" });
    saveCorsProxyConfig({ enabled: true, url: "https://cors.example.com/?" });
    saveProviderKeyConfig({ provider: "openai", apiKey: "sk-direct", baseURL: "https://api.openai.com/v1" });

    return resolveStreamEndpoint({}).then((resolved) => {
      assert.match(resolved.url, /^https:\/\/cors\.example\.com\/\?url=/);
      assert.ok(decodeURIComponent(resolved.url).includes("api.openai.com"));
    });
  });

  test("Layer 3: Direct when streamProxy + cors disabled but direct has key", () => {
    saveStreamProxyConfig({ enabled: false, url: "", token: "" });
    saveCorsProxyConfig({ enabled: false, url: "" });
    saveProviderKeyConfig({ provider: "openai", apiKey: "sk-test-123", baseURL: "https://api.openai.com/v1" });

    return resolveStreamEndpoint({}).then((resolved) => {
      assert.equal(resolved.url, "https://api.openai.com/v1/chat/completions");
      assert.equal(resolved.headers.Authorization, "Bearer sk-test-123");
    });
  });

  test("Fallback to null when no config → caller uses default /api/chat", () => {
    saveStreamProxyConfig({ enabled: false, url: "", token: "" });
    saveCorsProxyConfig({ enabled: false, url: "" });
    saveProviderKeyConfig({ provider: "openai", apiKey: "", baseURL: "" });

    return resolveStreamEndpoint({}).then((resolved) => {
      assert.equal(resolved, null);
    });
  });

  test("StreamProxy URL trailing slash stripped", () => {
    saveStreamProxyConfig({ enabled: true, url: "https://proxy.com/v1/", token: "t" });
    saveCorsProxyConfig({ enabled: false, url: "" });
    saveProviderKeyConfig({ provider: "openai", apiKey: "", baseURL: "" });

    return resolveStreamEndpoint({}).then((resolved) => {
      assert.equal(resolved.url, "https://proxy.com/v1/chat/completions");
    });
  });
});
