/**
 * Tool Coverage Map
 * Pattern từ office-coding-agent/scripts/check-tool-coverage.ts
 *
 * Quét word-bridge.js → list tất cả public methods.
 * Ánh xạ tới word-tool-configs → biết method nào đã có tool config, method nào chưa.
 *
 * Chạy: `node scripts/tool-coverage-map.js`
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");

async function main() {
  // 1. Load word-tool-configs
  const configsPath = path.join(rootDir, "public/modules/word-tool-configs/index.js");
  const { WORD_TOOL_CONFIGS, executeWordTool } = await import(pathToFileURL(configsPath).href);

  // 2. Parse word-bridge.js — tìm tất cả public methods
  const wordBridgePath = path.join(rootDir, "public/modules/word-bridge.js");
  const source = fs.readFileSync(wordBridgePath, "utf-8");

  // Regex: "  async methodName(" hoặc "  methodName("
  const methodRegex = /^\s+(async\s+)?([a-z][a-zA-Z0-9_]*)\s*\(/gm;
  const methodNames = new Set();
  // Lọc bỏ các tên JavaScript built-in / Promise / utilities (false positive)
  const JS_BUILTINS = new Set([
    "constructor", "if", "for", "while", "return", "switch", "catch", "throw",
    "clearTimeout", "setTimeout", "reject", "resolve", "then", "catch",
    "console", "Promise", "Array", "Object", "String", "Number", "JSON", "Math",
    "Date", "Error", "Map", "Set", "Symbol",
  ]);
  for (const m of source.matchAll(methodRegex)) {
    const name = m[2];
    if (name.startsWith("_")) continue;
    if (JS_BUILTINS.has(name)) continue;
    methodNames.add(name);
  }

  // 3. Lấy tên tools trong configs
  const configTools = new Set(WORD_TOOL_CONFIGS.map((c) => c.name));
  // Map: tool config nào cover method nào (heuristic đơn giản theo tên)
  function methodToToolMap(methodName) {
    const m = methodName.toLowerCase();
    if (m.startsWith("get") && m.includes("selection")) return ["word_read_selection"];
    if (m.startsWith("get") && (m.includes("doc") || m.includes("full"))) return ["word_read_document"];
    if (m.startsWith("get") && m.includes("paragraph")) return ["word_read_paragraphs"];
    if (m === "searchbody" || m.startsWith("search")) return ["word_search_text"];
    if (m === "replaceselection" || m === "insertatcursor" || m === "insertparagraphafter" || m === "insertafteranchor" || m === "appenddocumentend") return ["word_insert_text", "word_insert_markdown"];
    if (m.includes("comment") || m.includes("reply") || m.includes("resolve")) return ["word_insert_comment", "word_accept_comment", "word_resolve_comment", "word_get_tracked_changes"];
    if (m.includes("track") || m.includes("metadata") || m.includes("trackchanges")) return ["word_insert_with_track_changes", "word_get_tracked_changes"];
    if (m === "executetool" || m === "withtrackchanges" || m === "isavailable" || m === "gettextanalytics") return [];
    return [];
  }

  // 4. Build report
  const rows = [];
  for (const method of [...methodNames].sort()) {
    const tools = methodToToolMap(method);
    const covered = tools.length > 0;
    rows.push({ method, tools, covered });
  }

  console.log("\n=== Ori AI — Tool Coverage Map ===\n");
  console.log(`word-bridge.js public methods: ${methodNames.size}`);
  console.log(`word-tool-configs tools:       ${configTools.size}\n`);

  console.log("Method                         | Tool Configs");
  console.log("-------------------------------|-----------------------------------");
  for (const r of rows) {
    const status = r.covered ? "✅" : "⚠️ ";
    const tools = r.tools.length ? r.tools.join(", ") : "(utility — no MCP tool needed)";
    console.log(`${status} ${r.method.padEnd(29)} | ${tools}`);
  }

  console.log("\n=== Tool Configs not yet exposed via word-bridge ===\n");
  const exposedMethods = new Set();
  for (const r of rows) for (const t of r.tools) exposedMethods.add(t);
  for (const c of WORD_TOOL_CONFIGS) {
    if (!exposedMethods.has(c.name)) {
      console.log(`  • ${c.name} (config exists, no word-bridge method exposes it yet)`);
    }
  }
  console.log("");
}

main().catch((err) => {
  console.error("tool-coverage-map failed:", err);
  process.exit(1);
});
