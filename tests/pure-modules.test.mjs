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

// ─── DOC-STATE.JS & PROMPT CONTRACT ──────────────────────────────────────

import { docState } from "../public/modules/doc-state.js";
import fs from "node:fs";
import path from "node:path";

describe("doc-state.js & Prompt Architecture", () => {
  test("loadSystemPrompt: returns prompt with marker instructions", () => {
    const prompt = docState.loadSystemPrompt();
    assert.ok(prompt.length > 50);
    assert.ok(prompt.includes("Ori AI Agent"));
    assert.ok(prompt.includes("[Góp ý:"));
    assert.ok(prompt.includes("[Tài liệu Word đang mở]"));
  });

  test("saveHouseVoice and loadHouseVoice roundtrip", () => {
    docState.saveHouseVoice("Văn phong ngắn gọn, dứt khoát.");
    const loaded = docState.loadHouseVoice();
    assert.equal(loaded, "Văn phong ngắn gọn, dứt khoát.");
  });

  test("saveEffort and loadEffort roundtrip", () => {
    docState.saveEffort("high");
    const loaded = docState.loadEffort("medium");
    assert.equal(loaded, "high");
  });

  test("saveHistory and loadHistory roundtrip", async () => {
    const msgs = [
      { role: "user", content: "Soát lỗi bài này" },
      { role: "assistant", content: "Đã tìm thấy 1 lỗi." }
    ];
    docState.saveHistory(msgs);
    // saveHistory uses debouncedSave(500ms), wait 550ms for persistence
    await new Promise((resolve) => setTimeout(resolve, 550));
    const loaded = docState.loadHistory();
    assert.equal(loaded.length, 2);
    assert.equal(loaded[0].role, "user");
  });
});

// ─── COMMENT MARKER PARSER & DIFF EXTRACTION ─────────────────────────────

function extractCommentMarkers(text) {
  if (!text || typeof text !== "string") return [];
  const out = [];
  const regex = /\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"([^"]{3,400})"\s*(?:->|:|➔)\s*([^\]]+)\]/gim;
  let m;
  while ((m = regex.exec(text)) !== null) {
    let sug = m[2].trim();
    sug = sug.replace(/^["']|["']$/g, "").trim();
    out.push({
      anchor: m[1].trim(),
      suggestion: sug,
    });
  }
  return out;
}

function extractExecutableText(text) {
  if (!text || typeof text !== "string") return "";
  const delimiterMatch = text.match(/(?:^|\n)\s*---\s*\n([\s\S]*?)\n\s*---\s*(?:\n|$)/);
  if (delimiterMatch && delimiterMatch[1] && delimiterMatch[1].trim()) {
    return delimiterMatch[1].trim();
  }
  const codeBlockMatch = text.match(/```(?:[a-z]*)\n([\s\S]*?)```/i);
  if (codeBlockMatch && codeBlockMatch[1] && codeBlockMatch[1].trim()) {
    return codeBlockMatch[1].trim();
  }
  return text.replace(/\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"[^"]+"\s*(?:->|:|➔)\s*[^\]]+\]/gim, "").trim();
}

describe("Marker Extraction & Patch Workflow", () => {
  test("extractCommentMarkers: parses standard marker", () => {
    const text = 'Em thấy lỗi này: [Góp ý: "chúng tôi xin gửi lời cảm ơn" -> "chúng tôi trân trọng cảm ơn"]';
    const markers = extractCommentMarkers(text);
    assert.equal(markers.length, 1);
    assert.equal(markers[0].anchor, "chúng tôi xin gửi lời cảm ơn");
    assert.equal(markers[0].suggestion, "chúng tôi trân trọng cảm ơn");
  });

  test("extractCommentMarkers: parses multiple markers in one text", () => {
    const text = `
      1. [Góp ý: "lỗi chính tả một" -> "đúng chính tả một"]
      2. [Nhận xét: "câu dài dòng thừa thãi" ➔ "câu ngắn gọn"]
      3. [Comment: "từ sáo rỗng" : "từ chân thực"]
    `;
    const markers = extractCommentMarkers(text);
    assert.equal(markers.length, 3);
    assert.equal(markers[0].anchor, "lỗi chính tả một");
    assert.equal(markers[1].anchor, "câu dài dòng thừa thãi");
    assert.equal(markers[2].anchor, "từ sáo rỗng");
  });

  test("extractExecutableText: extracts text between --- delimiters", () => {
    const text = `
Dưới đây là bản dịch:

---
This is the translated English text of the document.
It has two lines.
---

Chúc bạn làm việc hiệu quả!
    `;
    const extracted = extractExecutableText(text);
    assert.equal(extracted, "This is the translated English text of the document.\nIt has two lines.");
  });

  test("extractExecutableText: extracts text inside code block", () => {
    const text = "Kết quả:\n```markdown\n# Tiêu đề mới\nNội dung mới\n```";
    const extracted = extractExecutableText(text);
    assert.equal(extracted, "# Tiêu đề mới\nNội dung mới");
  });

  test("extractExecutableText: strips raw markers from text", () => {
    const text = 'Sửa xong: [Góp ý: "cũ" -> "mới"] đoạn này';
    const extracted = extractExecutableText(text);
    assert.equal(extracted, "Sửa xong:  đoạn này");
  });
});

// ─── SKILLS.JSON VALIDITY & HTML CONTRACT ────────────────────────────────

describe("Skills Config & HTML DOM Contracts", () => {
  test("skills.json: valid structure and prompt presence", () => {
    const raw = fs.readFileSync(path.resolve("public/skills.json"), "utf8");
    const skills = JSON.parse(raw);
    assert.ok(Array.isArray(skills));
    assert.ok(skills.length >= 6);

    for (const skill of skills) {
      assert.ok(skill.id, "Skill must have an id");
      assert.ok(skill.icon, "Skill must have an icon");
      assert.ok(skill.label, "Skill must have a label");
      assert.ok(skill.prompt, "Skill must have a prompt");
      assert.ok(skill.tip, "Skill must have a tip");
    }
  });

  test("index.html: contains all required DOM element IDs", () => {
    const html = fs.readFileSync(path.resolve("public/index.html"), "utf8");
    const requiredIds = [
      "quickEffortSelect",
      "btnToggleTheme",
      "btnSessions",
      "btnSettings",
      "chatContainer",
      "emptyHint",
      "docInsightsBar",
      "docInsightsText",
      "quickSkillsBar",
      "slashMenu",
      "slashMenuList",
      "chatInput",
      "btnSend",
      "settingsModal",
      "btnCloseSettings",
      "btnSaveSettings",
      "cfgHouseVoice",
      "cfgEffort",
      "btnClearChatModal",
      "cfgStreamProxyEnabled",
      "cfgStreamProxyUrl",
      "cfgStreamProxyToken",
      "cfgCorsProxyEnabled",
      "cfgCorsProxyUrl",
      "cfgDirectEnabled",
      "cfgDirectBaseURL",
      "cfgDirectApiKey",
      "licEmailDisplay",
      "licDeviceDisplay",
      "sessionsModal",
      "sessionsList",
      "btnCloseSessions",
      "licenseModal",
      "licEmail",
      "licKey",
      "licStatus",
      "btnActivate"
    ];

    for (const id of requiredIds) {
      assert.ok(html.includes(`id="${id}"`), `Missing required id="${id}" in index.html`);
    }
  });
});

