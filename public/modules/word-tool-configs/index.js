/**
 * Word Tool Configs (declarative pattern)
 *
 * Mỗi tool = { name, description, params, execute(ctx, args) }
 * - params: JSON Schema mô tả
 * - execute: chạy trong Word.run context của Office.js
 *
 * Pattern từ office-coding-agent/src/tools/codegen/.
 * Tách khỏi word-bridge.js để 1 nguồn truth cho cả UI + MCP server.
 *
 * Dùng cho:
 *  - MCP server (server/tools.js) — sinh schema tự động từ configs
 *  - UI buttons trong app.js — gọi trực tiếp config.execute(ctx, args)
 *  - E2E test — gọi trong Word thật qua config.execute()
 */

import { wordBridge } from "../word-bridge.js?v=2.4.0";

/**
 * Factory chạy 1 tool config trong 1 Word.run block.
 * Mỗi tool.execute(context, args) được gọi với Office.js context đã sẵn sàng.
 */
async function runInWord(config, args) {
  if (!wordBridge.isAvailable()) {
    throw new Error("Office.js không khả dụng — chỉ chạy trong Microsoft Word.");
  }
  return Word.run(async (context) => {
    try {
      const result = await config.execute(context, args);
      await context.sync();
      return result;
    } catch (err) {
      // Trả về string thay vì throw để caller (agent loop) không vỡ
      return { ok: false, error: err.message || String(err) };
    }
  });
}

// ─── TOOL: read_selection ─────────────────────────────────────────────────

const readSelection = {
  name: "word_read_selection",
  description: "Đọc đoạn văn bản user đang bôi đen trong Word (text thuần hoặc GFM Markdown).",
  params: {
    type: "object",
    properties: {
      format: { type: "string", enum: ["text", "markdown"], default: "text" },
    },
  },
  execute: async (context, { format = "text" } = {}) => {
    if (format === "markdown") {
      return await wordBridge.getSelectionAsMarkdown();
    }
    return { text: await wordBridge.getSelectedText() };
  },
};

// ─── TOOL: read_document ─────────────────────────────────────────────────

const readDocument = {
  name: "word_read_document",
  description: "Đọc toàn bộ văn bản tài liệu Word. Trả về plain text (1 sync, nhanh).",
  params: {
    type: "object",
    properties: {
      limit: { type: "integer", default: 30000, description: "Giới hạn ký tự" },
    },
  },
  execute: async (context, { limit = 30000 } = {}) => {
    return { text: await wordBridge.getFullDocumentText(limit) };
  },
};

// ─── TOOL: read_paragraphs ──────────────────────────────────────────────

const readParagraphs = {
  name: "word_read_paragraphs",
  description: "Đọc danh sách đoạn văn có cấu trúc (style, text).",
  params: {
    type: "object",
    properties: {
      styleFilter: { type: "string", description: "Lọc theo style (Heading 1, Normal...)" },
      limit: { type: "integer", default: 0, description: "Số đoạn tối đa" },
      textLimit: { type: "integer", default: 300 },
    },
  },
  execute: async (context, args) => {
    return await wordBridge.getParagraphs(args.styleFilter, args.limit, args.textLimit);
  },
};

// ─── TOOL: search_text ──────────────────────────────────────────────────

const searchText = {
  name: "word_search_text",
  description:
    "Tìm text trong document body. Trả về search-anchor contract (anchor_search_text, anchor_match_index, anchor_search_options) để insertAfterAnchor / insertMarkdown dùng chính xác.",
  params: {
    type: "object",
    properties: {
      searchText: { type: "string" },
      matchCase: { type: "boolean", default: false },
      matchWholeWord: { type: "boolean", default: false },
      matchIndex: { type: "integer", default: 0, description: "Match thứ mấy (0-based)" },
    },
    required: ["searchText"],
  },
  execute: async (context, args) => {
    return await wordBridge.searchBody(args.searchText, {
      matchCase: args.matchCase,
      matchWholeWord: args.matchWholeWord,
      matchIndex: args.matchIndex || 0,
    });
  },
};

// ─── TOOL: insert_text ──────────────────────────────────────────────────

const insertText = {
  name: "word_insert_text",
  description: "Chèn text thuần tại cursor (Markdown sẽ bị strip). Dùng insertMarkdown nếu cần Rich Text.",
  params: {
    type: "object",
    properties: {
      text: { type: "string" },
      location: {
        type: "string",
        enum: ["replace", "after", "before", "end"],
        default: "replace",
      },
    },
    required: ["text"],
  },
  execute: async (context, { text, location = "replace" } = {}) => {
    if (location === "end") return await wordBridge.appendDocumentEnd(text);
    if (location === "after") return await wordBridge.insertParagraphAfter(text);
    return await wordBridge.replaceSelection(text);
  },
};

// ─── TOOL: insert_markdown (hero tool) ──────────────────────────────────

const insertMarkdown = {
  name: "word_insert_markdown",
  description:
    "Chèn Markdown vào Word (parse → HTML → DOMPurify sanitize → insertHtml). Giữ heading/list/strong/em/code/link/table. Hỗ trợ search-anchor: nếu anchor_search_text được cung cấp, insert tại vị trí match (echo lại từ searchText tool).",
  params: {
    type: "object",
    properties: {
      markdown: { type: "string" },
      where: {
        type: "string",
        enum: ["replace", "after", "before", "end"],
        default: "replace",
      },
      anchor_search_text: { type: "string", description: "Tìm anchor này thay vì dùng selection hiện tại" },
      anchor_match_index: { type: "integer", default: 0 },
      anchor_match_case: { type: "boolean", default: false },
      anchor_match_whole_word: { type: "boolean", default: false },
    },
    required: ["markdown"],
  },
  execute: async (context, args) => {
    const { markdown, where = "replace", anchor_search_text, anchor_match_index = 0, anchor_match_case = false, anchor_match_whole_word = false } = args;

    if (anchor_search_text) {
      const searchOptions = { matchCase: anchor_match_case, matchWholeWord: anchor_match_whole_word };
      const html = await import("../word-bridge.js?v=2.4.0").then((m) => m.markdownToHtml(markdown));
      await context.sync();
      const temp = document.createElement("div");
      temp.innerHTML = html;
      const text = temp.textContent || "";
      const result = await wordBridge.insertAfterAnchor(anchor_search_text, text, {
        matchIndex: anchor_match_index,
        searchOptions,
      });
      return {
        ok: true,
        where: "after_anchor",
        anchor_search_text,
        anchor_match_index,
        inserted: text,
      };
    }

    if (where === "end") {
      const html = await import("../word-bridge.js?v=2.4.0").then((m) => m.markdownToHtml(markdown));
      const sel = context.document.body;
      sel.insertHtml(html, Word.InsertLocation.end);
      await context.sync();
      return { ok: true, where: "end" };
    }
    if (where === "replace") {
      await wordBridge.replaceSelection(markdown);
      return { ok: true, where: "replace" };
    }
    if (where === "after") {
      await wordBridge.insertParagraphAfter(markdown);
      return { ok: true, where: "after" };
    }
    return { ok: false, error: "before_selection chưa được hỗ trợ — dùng anchor_search_text thay thế" };
  },
};

// ─── TOOL: insert_comment (Ori Agent prefix) ──────────────────────────────

const insertComment = {
  name: "word_insert_comment",
  description: "Chèn 1 margin comment ở lề Word với prefix '[Ori Agent]'. Comment có thể Accept (replace text + xoá) hoặc Reject (xoá).",
  params: {
    type: "object",
    properties: {
      anchorText: { type: "string" },
      commentText: { type: "string" },
    },
    required: ["anchorText", "commentText"],
  },
  execute: async (context, args) => {
    return await wordBridge.insertComment(args.commentText, args.anchorText, { authorTag: true });
  },
};

// ─── TOOL: accept_comment ───────────────────────────────────────────────

const acceptComment = {
  name: "word_accept_comment",
  description: "Áp dụng suggestion từ comment: thay text bằng replacementText rồi xoá comment. (Khi user bấm ✓ trên panel.)",
  params: {
    type: "object",
    properties: {
      anchorText: { type: "string" },
      replacementText: { type: "string" },
    },
    required: ["anchorText", "replacementText"],
  },
  execute: async (context, args) => {
    return await wordBridge.acceptComment({
      anchorText: args.anchorText,
      replacementText: args.replacementText,
    });
  },
};

// ─── TOOL: resolve_comment ──────────────────────────────────────────────

const resolveComment = {
  name: "word_resolve_comment",
  description: "Đánh dấu 1 comment là đã giải quyết (Word.js 1.4+). Comment vẫn còn trong lề.",
  params: {
    type: "object",
    properties: {
      anchorText: { type: "string" },
    },
    required: ["anchorText"],
  },
  execute: async (context, args) => {
    return await wordBridge.resolveComment({ anchorText: args.anchorText });
  },
};

// ─── TOOL: insert_with_track_changes ───────────────────────────────────

const insertWithTrackChanges = {
  name: "word_insert_with_track_changes",
  description: "Chèn text với Track Changes (Redlines). Bật changeTrackingMode, insert, restore mode cũ.",
  params: {
    type: "object",
    properties: {
      text: { type: "string" },
      location: { type: "string", enum: ["replace", "after", "before"], default: "replace" },
    },
    required: ["text"],
  },
  execute: async (context, args) => {
    return await wordBridge.insertWithTrackChanges(args.text, args.text, "Ori AI");
  },
};

// ─── TOOL: get_tracked_changes ─────────────────────────────────────────

const getTrackedChanges = {
  name: "word_get_tracked_changes",
  description: "Liệt kê các tracked changes đang chờ (3s timeout tránh Mac Word.js hang).",
  params: {
    type: "object",
    properties: {
      timeoutMs: { type: "integer", default: 3000 },
    },
  },
  execute: async (context, args) => {
    return await wordBridge.getTrackedChanges(args.timeoutMs || 3000);
  },
};

// ─── 6 tools mới fill coverage gap ─────────────────────────────────────

const findReplace = {
  name: "word_find_replace",
  description: "Tìm và thay thế chuỗi trong toàn bộ tài liệu Word.",
  params: {
    type: "object",
    properties: {
      find: { type: "string" },
      replace: { type: "string" },
      matchCase: { type: "boolean", default: false },
      matchWholeWord: { type: "boolean", default: false },
      maxReplacements: { type: "integer", default: 0, description: "0 = tất cả" },
    },
    required: ["find", "replace"],
  },
  execute: async (context, args) => {
    return await wordBridge.findReplace(
      args.find,
      args.replace,
      args.matchCase,
      args.matchWholeWord,
      args.maxReplacements
    );
  },
};

const deleteText = {
  name: "word_delete_text",
  description: "Tìm và xoá text trong tài liệu Word.",
  params: {
    type: "object",
    properties: {
      find: { type: "string" },
      maxDeletions: { type: "integer", default: 0, description: "0 = tất cả" },
    },
    required: ["find"],
  },
  execute: async (context, args) => {
    return await wordBridge.deleteText(args.find, args.maxDeletions);
  },
};

const setParagraphStyle = {
  name: "word_set_paragraph_style",
  description: "Gán style cho đoạn văn chứa anchor (Heading 1, Heading 2, Normal, Quote...).",
  params: {
    type: "object",
    properties: {
      anchor: { type: "string" },
      style: { type: "string", description: "Tên style (Heading 1, Normal, Quote, Title...)" },
    },
    required: ["anchor", "style"],
  },
  execute: async (context, args) => {
    return await wordBridge.setParagraphStyle(args.anchor, args.style);
  },
};

const selectRange = {
  name: "word_select_range",
  description: "Tìm anchor trong Word và bôi đen (highlight) đoạn đó — dùng cho click-to-jump.",
  params: {
    type: "object",
    properties: {
      anchor: { type: "string" },
      paragraphIndex: { type: "integer", description: "Optional. Nếu có dùng index, không cần anchor." },
    },
  },
  execute: async (context, args) => {
    return await wordBridge.selectRange(args.anchor, args.paragraphIndex);
  },
};

const insertOoxml = {
  name: "word_insert_ooxml",
  description: "Chèn OOXML fragment tùy biến tại anchor. OOXML phải là fragment <w:p>...</w:p>, KHÔNG phải full package.",
  params: {
    type: "object",
    properties: {
      anchor: { type: "string" },
      ooxml: { type: "string" },
      location: { type: "string", enum: ["after", "before", "replace"], default: "after" },
    },
    required: ["anchor", "ooxml"],
  },
  execute: async (context, args) => {
    return await wordBridge.insertOoxml(args.anchor, args.ooxml, args.location);
  },
};

const getTextAnalytics = {
  name: "word_get_text_analytics",
  description: "Đếm số từ, số ký tự, thời gian đọc ước tính của 1 đoạn text.",
  params: {
    type: "object",
    properties: {
      text: { type: "string" },
    },
    required: ["text"],
  },
  execute: async (context, args) => {
    return wordBridge.getTextAnalytics(args.text);
  },
};

// ─── Registry ────────────────────────────────────────────────────────────

export const WORD_TOOL_CONFIGS = [
  readSelection,
  readDocument,
  readParagraphs,
  searchText,
  insertText,
  insertMarkdown,
  insertComment,
  acceptComment,
  resolveComment,
  insertWithTrackChanges,
  getTrackedChanges,
  findReplace,
  deleteText,
  setParagraphStyle,
  selectRange,
  insertOoxml,
  getTextAnalytics,
];

/**
 * Helper: tìm config theo name.
 */
export function getWordToolConfig(name) {
  return WORD_TOOL_CONFIGS.find((c) => c.name === name) || null;
}

/**
 * Helper: chạy 1 tool theo name (dùng cho UI button handler).
 */
export async function executeWordTool(name, args) {
  const config = getWordToolConfig(name);
  if (!config) throw new Error(`Tool không tồn tại: ${name}`);
  return await runInWord(config, args);
}
