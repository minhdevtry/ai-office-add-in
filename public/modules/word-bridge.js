/**
 * Word Bridge Module
 * Toàn bộ logic tương tác trực tiếp với Microsoft Word qua Office.js
 *
 * Pattern tham khảo:
 *  - vivword: dùng setSelectedDataAsync cho mọi insert/replace (ổn định nhất)
 *  - claude-word-addin: Track Changes = toggle changeTrackingMode + insert (Word tự mark)
 *  - wordbridge: serializeError chi tiết + Promise.race timeout cho Mac Word.js
 *  - pi-for-word: marked + DOMPurify cho markdown→HTML, search-anchor contract
 *  - office-coding-agent: declarative tool config {name, description, params, execute}
 *  - draftspect: withTrackChanges(context, bool, body) wrapper
 */

/* global Office, Word, marked, DOMPurify */

// marked + DOMPurify load từ CDN (index.html), KHÔNG dùng ESM import vì
// Office iframe không support bare specifier (browser không resolve).
const _marked = typeof marked !== "undefined" ? marked : null;
const _DOMPurify = typeof DOMPurify !== "undefined" ? DOMPurify : null;

import { computeWordDiff } from "./diff.js?v=2.4.0";

// ─── Constants ─────────────────────────────────────────────────────────────

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const GET_TRACKED_CHANGES_TIMEOUT_MS = 3000;
const SET_TRACK_CHANGES_TIMEOUT_MS = 5000;
const GET_DOC_METADATA_TIMEOUT_MS = 4000;

// DOMPurify allowlist — bổ sung `input` cho GFM task list, `checked` cho checkbox
const HTML_SANITIZE_OPTIONS = {
  USE_PROFILES: { html: true },
  ADD_TAGS: ["input"],
  ADD_ATTR: ["checked", "disabled", "type"],
};

// ─── Utilities ─────────────────────────────────────────────────────────────

/**
 * Serialize lỗi Office.js đầy đủ để debug (copy từ wordbridge).
 * Pull cả debugInfo.code, errorLocation, statement, surroundingStatements.
 */
export function serializeError(err) {
  if (!err) return "null";
  const out = {};
  for (const k of ["name", "code", "message", "stack"]) {
    if (err[k] !== undefined) out[k] = err[k];
  }
  if (err.debugInfo) {
    out.debugInfo = {};
    for (const k of [
      "code",
      "message",
      "errorLocation",
      "statement",
      "surroundingStatements",
      "fullStatements",
    ]) {
      if (err.debugInfo[k] !== undefined) out.debugInfo[k] = err.debugInfo[k];
    }
  }
  try {
    return JSON.stringify(out, null, 2);
  } catch {
    return String(err);
  }
}

/**
 * Run một async function với timeout guard.
 * Trả về { ok: true, value } hoặc { ok: false, error, timedOut }.
 */
async function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(
      () => reject(new Error(`${label} timed out after ${ms}ms`)),
      ms
    );
  });
  try {
    const value = await Promise.race([promise, timeout]);
    return { ok: true, value };
  } catch (err) {
    return { ok: false, error: err, timedOut: err.message?.includes("timed out") };
  } finally {
    clearTimeout(timer);
  }
}

export function escapeXml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Chuyển Markdown sang HTML sạch để Word parse thành Rich Text.
 *
 * Dùng `marked` (parse GFM đầy đủ: heading, list, table, code block, task list, link)
 * + `DOMPurify` để sanitize chống XSS khi insert vào Word.
 *
 * Pattern từ pi-for-word/word-tool-insert-markdown.js.
 *
 * Lưu ý: marked + DOMPurify load từ CDN (index.html). Trong môi trường dev/test
 * (Node.js), sẽ fallback về regex tự viết.
 */
export function markdownToHtml(md) {
  if (!md) return "";
  const trimmed = String(md).trim();
  if (!trimmed) return "";

  if (!_marked || !_DOMPurify) {
    // Fallback: regex đơn giản (chỉ heading + bold + italic + list)
    return simpleMarkdownFallback(trimmed);
  }

  const raw = _marked.parse(trimmed, { async: false, gfm: true, breaks: true });
  if (!raw || typeof raw !== "string" || !raw.trim()) return "";

  return _DOMPurify.sanitize(raw, HTML_SANITIZE_OPTIONS);
}

/**
 * Fallback regex đơn giản khi marked + DOMPurify chưa load (Node test, dev mode).
 */
function simpleMarkdownFallback(md) {
  let html = String(md)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
  html = html.replace(/^## (.*$)/gim, "<h2>$1</h2>");
  html = html.replace(/^# (.*$)/gim, "<h1>$1</h1>");
  html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/\*(.*?)\*/g, "<em>$1</em>");
  html = html.replace(/^[-*] (.*$)/gim, "<li>$1</li>");
  html = html.replace(/(<li>.*<\/li>\n?)+/g, (m) => `<ul>${m}</ul>`);
  html = html.replace(/\n{2,}/g, "</p><p>");
  return `<p>${html}</p>`.replace(/<p><(h\d|ul|li)/g, "<$1").replace(/<\/(h\d|ul|li)><\/p>/g, "</$1>");
}

/**
 * Chuyển HTML (từ `range.getHtml()` của Word) sang GFM Markdown.
 * Helper nội bộ cho `getSelectionAsMarkdown()`.
 *
 * Word wrap selection trong <p> / <h1-9> / <ul> / <ol> / <li> / <strong> / <em> / <a>.
 */
function htmlToMarkdown(html) {
  if (!html || typeof html !== "string") return "";
  let s = html;

  // Word thường prefix XML declaration & namespaces — strip
  s = s.replace(/^<\?xml[^>]*>/, "").trim();

  // H1-H6
  s = s.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, (_, t) => `# ${stripTags(t).trim()}\n\n`);
  s = s.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, (_, t) => `## ${stripTags(t).trim()}\n\n`);
  s = s.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, (_, t) => `### ${stripTags(t).trim()}\n\n`);
  s = s.replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, (_, t) => `#### ${stripTags(t).trim()}\n\n`);
  s = s.replace(/<h5[^>]*>([\s\S]*?)<\/h5>/gi, (_, t) => `##### ${stripTags(t).trim()}\n\n`);
  s = s.replace(/<h6[^>]*>([\s\S]*?)<\/h6>/gi, (_, t) => `###### ${stripTags(t).trim()}\n\n`);

  // Lists
  s = s.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (_, list) => {
    return list.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (_, item) => `- ${stripTags(item).trim()}\n`).trim() + "\n\n";
  });
  s = s.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (_, list) => {
    let i = 0;
    return list.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (_, item) => `${++i}. ${stripTags(item).trim()}\n`).trim() + "\n\n";
  });

  // Inline
  s = s.replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, "**$1**");
  s = s.replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, "**$1**");
  s = s.replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, "*$1*");
  s = s.replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, "*$1*");
  s = s.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, "`$1`");
  s = s.replace(/<a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, "[$2]($1)");

  // Block-level: <p>, <div>, <br>
  s = s.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, (_, t) => `${stripTags(t).trim()}\n\n`);
  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/div>/gi, "\n");

  // Strip remaining tags
  s = stripTags(s);

  // Cleanup whitespace
  s = s.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

function stripTags(s) {
  return String(s || "").replace(/<[^>]+>/g, "");
}

// ─── WordBridge Class ───────────────────────────────────────────────────────

export class WordBridge {
  constructor() {
    this._trackChangesSupported = null;
    this._commentsSupported = null;
    this.cachedSelection = "";
  }

  isAvailable() {
    return typeof Office !== "undefined" && typeof Word !== "undefined";
  }

  get commentsSupported() {
    if (this._commentsSupported === null) {
      try {
        this._commentsSupported =
          Office.context.requirements.isSetSupported("WordApi", "1.4");
      } catch {
        this._commentsSupported = false;
      }
    }
    return this._commentsSupported;
  }

  get trackChangesSupported() {
    if (this._trackChangesSupported === null) {
      try {
        this._trackChangesSupported =
          Office.context.requirements.isSetSupported("WordApi", "1.4");
      } catch {
        this._trackChangesSupported = false;
      }
    }
    return this._trackChangesSupported;
  }

  // ─── READ ──────────────────────────────────────────────────────────────

  /**
   * Đọc đoạn văn bản người dùng đang bôi đen (text thuần).
   * Ưu tiên Common API (ổn định hơn Word.run trên iPad/Web).
   */
  async getSelectedText() {
    if (!this.isAvailable()) return this.cachedSelection || "";
    return new Promise((resolve, reject) => {
      try {
        Office.context.document.getSelectedDataAsync(
          Office.CoercionType.Text,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const text = (result.value || "").trim();
              this.cachedSelection = text;
              resolve(text);
            } else {
              Word.run(async (context) => {
                const sel = context.document.getSelection();
                sel.load("text");
                await context.sync();
                const text = (sel.text || "").trim();
                this.cachedSelection = text;
                resolve(text);
              }).catch((e) => reject(e));
            }
          }
        );
      } catch (err) {
        reject(err);
      }
    });
  }

  /**
   * Đọc selection dưới dạng GFM Markdown (HTML → Markdown).
   * Hữu ích cho AI agent khi cần hiểu cấu trúc selection (bold, heading, list, link)
   * thay vì plain text mất formatting.
   *
   * Pattern từ pi-for-word/word-tool-get-selection.js: dùng `range.getHtml()` rồi convert HTML → GFM.
   *
   * @returns {Promise<{ text: string, markdown: string, html: string }>}
   */
  async getSelectionAsMarkdown() {
    if (!this.isAvailable()) return { text: this.cachedSelection || "", markdown: "", html: "" };
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load(["text", "html"]);
      await context.sync();
      const text = (sel.text || "").trim();
      const html = sel.html || "";
      // Convert HTML → GFM Markdown (giữ heading/list/strong/em/code)
      const markdown = htmlToMarkdown(html);
      this.cachedSelection = text;
      return { text, markdown, html };
    });
  }

  /**
   * Đọc toàn bộ văn bản tài liệu Word.
   * Dùng `body.text` (1 sync) thay vì loop paragraphs.items[] (N sync calls).
   */
  async getFullDocumentText(limit = 0) {
    if (!this.isAvailable()) return "";
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      const full = body.text || "";
      return limit > 0 && full.length > limit ? full.slice(0, limit) : full;
    });
  }

  /**
   * Đọc danh sách đoạn văn có cấu trúc.
   */
  async getParagraphs(styleFilter = null, limit = 0, textLimit = 300) {
    if (!this.isAvailable()) return { total: 0, returned: 0, paragraphs: [] };
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load(["text", "style", "styleBuiltIn"]);
      await context.sync();
      const items = [];
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const text = p.text || "";
        if (text.trim() === "") continue;
        if (styleFilter && p.style !== styleFilter && p.styleBuiltIn !== styleFilter) continue;
        items.push({
          index: i,
          style: p.style,
          styleBuiltIn: p.styleBuiltIn,
          text: text.length > textLimit ? text.slice(0, textLimit) + "…" : text,
          length: text.length,
        });
        if (limit > 0 && items.length >= limit) break;
      }
      return {
        total: paragraphs.items.length,
        returned: items.length,
        paragraphs: items,
      };
    });
  }

  /**
   * Thuật toán 3 lớp định vị và bôi đen đoạn văn bản (Click-to-Jump).
   */
  async selectRange(anchor, paragraphIndex = null) {
    if (!this.isAvailable()) return { ok: false, error: "Office.js không khả dụng" };
    return Word.run(async (context) => {
      if (anchor && typeof anchor === "string" && anchor.trim()) {
        const cleanAnchor = anchor.trim();
        let searchResults = context.document.body.search(cleanAnchor, { matchCase: false });
        searchResults.load("items");
        await context.sync();
        if (searchResults.items.length > 0) {
          searchResults.items[0].select();
          await context.sync();
          return { ok: true, matched: "exact", count: searchResults.items.length };
        }
        if (cleanAnchor.length > 25) {
          const prefix = cleanAnchor.slice(0, 25);
          searchResults = context.document.body.search(prefix, { matchCase: false });
          searchResults.load("items");
          await context.sync();
          if (searchResults.items.length > 0) {
            searchResults.items[0].select();
            await context.sync();
            return { ok: true, matched: "fuzzy_prefix", count: searchResults.items.length };
          }
        }
      }
      if (paragraphIndex !== null && paragraphIndex !== undefined && paragraphIndex >= 0) {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        if (paragraphIndex < paragraphs.items.length) {
          paragraphs.items[paragraphIndex].select();
          await context.sync();
          return { ok: true, matched: "paragraph_index", index: paragraphIndex };
        }
      }
      throw new Error(`Không tìm thấy vị trí neo: '${anchor || paragraphIndex}'`);
    });
  }

  // ─── COMMENTS ──────────────────────────────────────────────────────────

  /**
   * Chèn bình luận mép lề phải (Margin Comments via WordApi 1.4).
   * Comment từ AI Agent tự động prefix "[Ori Agent] " để user phân biệt.
   */
  async insertComment(commentText, anchor = null, { authorTag = true } = {}) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");
    if (!this.commentsSupported || typeof Word?.Comment === "undefined") {
      throw new Error(
        "Phiên bản Word hiện tại không hỗ trợ chèn bình luận (cần Word 2016 trở lên hoặc Microsoft 365)."
      );
    }
    const finalText = authorTag ? `[Ori Agent] ${commentText}` : commentText;
    return Word.run(async (context) => {
      let targetRange;
      if (anchor && typeof anchor === "string" && anchor.trim()) {
        const results = context.document.body.search(anchor.trim(), { matchCase: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
          targetRange = results.items[0];
        }
      }
      if (!targetRange) targetRange = context.document.getSelection();
      targetRange.insertComment(finalText);
      await context.sync();
      return { ok: true, type: "margin_comment", text: finalText };
    });
  }

  /**
   * Lấy danh sách tất cả comment trong tài liệu (id, anchor, text, author).
   * Dùng để panel "Nhận xét chờ duyệt" biết có bao nhiêu comment từ [Ori Agent].
   */
  async getComments() {
    if (!this.isAvailable()) return [];
    if (!this.commentsSupported || typeof Word?.Comment === "undefined") return [];

    return Word.run(async (context) => {
      const comments = context.document.body.comments;
      comments.load(["items/id", "items/content", "items/author"]);
      await context.sync();
      return comments.items.map((c, i) => ({
        index: i,
        id: c.id || String(i),
        content: c.content,
        author: c.author || "",
        isAgent: typeof c.content === "string" && c.content.startsWith("[Ori Agent]"),
      }));
    });
  }

  /**
   * Tìm comment theo anchor (chuỗi text) trong document.
   * Dùng khi user click Accept mà add-in chỉ biết anchor text (chưa track id).
   */
  async findCommentByAnchor(anchorText) {
    if (!this.isAvailable()) return null;
    return Word.run(async (context) => {
      const results = context.document.body.search(anchorText, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) return null;
      // Lấy comment trên range đó
      const rangeComments = results.items[0].comments;
      rangeComments.load("items");
      await context.sync();
      if (rangeComments.items.length === 0) return null;
      const c = rangeComments.items[0];
      return {
        id: c.id || "",
        content: c.content || "",
        author: c.author || "",
      };
    });
  }

  /**
   * Xoá 1 comment theo ID (hoặc theo anchor nếu không có id).
   * Add-in Word không expose id ổn định qua API 1.4 — ta xoá theo anchor.
   */
  async deleteComment({ anchorText, commentId } = {}) {
    if (!this.isAvailable()) return { ok: false };
    return Word.run(async (context) => {
      let targetComment = null;
      if (anchorText) {
        const results = context.document.body.search(anchorText, { matchCase: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
          const rangeComments = results.items[0].comments;
          rangeComments.load("items");
          await context.sync();
          if (rangeComments.items.length > 0) targetComment = rangeComments.items[0];
        }
      }
      if (!targetComment) return { ok: false, error: "Không tìm thấy comment" };
      targetComment.delete();
      await context.sync();
      return { ok: true };
    });
  }

  /**
   * Accept 1 comment: thay text trong range bằng replacementText, rồi xoá comment.
   * Đây là flow chính khi user bấm "✓ Áp dụng" bên cạnh mỗi [Ori Agent] comment.
   */
  async acceptComment({ anchorText, replacementText } = {}) {
    if (!this.isAvailable()) return { ok: false };
    if (!anchorText || !replacementText) {
      return { ok: false, error: "Thiếu anchorText hoặc replacementText" };
    }
    return Word.run(async (context) => {
      const results = context.document.body.search(anchorText, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) {
        return { ok: false, error: `Không tìm thấy anchor: "${anchorText.slice(0, 40)}..."` };
      }
      const targetRange = results.items[0];
      // 1. Xoá comment trên range này
      const rangeComments = targetRange.comments;
      rangeComments.load("items");
      await context.sync();
      const commentToDelete = rangeComments.items[0];
      // 2. Replace text
      targetRange.insertText(replacementText, Word.InsertLocation.replace);
      await context.sync();
      // 3. Xoá comment
      if (commentToDelete) {
        commentToDelete.delete();
        await context.sync();
      }
      return { ok: true };
    });
  }

  /**
   * Reply vào 1 comment (threaded comment).
   * Word hỗ trợ comment.replies.add(text).
   */
  async replyToComment({ anchorText, replyText } = {}) {
    if (!this.isAvailable()) return { ok: false };
    return Word.run(async (context) => {
      const results = context.document.body.search(anchorText, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) return { ok: false, error: "Anchor not found" };
      const rangeComments = results.items[0].comments;
      rangeComments.load("items/replies");
      await context.sync();
      if (rangeComments.items.length === 0) return { ok: false, error: "No comment" };
      rangeComments.items[0].replies.add(replyText);
      await context.sync();
      return { ok: true };
    });
  }

  /**
   * Resolve (đóng) 1 comment — đánh dấu "Đã giải quyết" trong Word.
   * Comment vẫn còn trong lề, chỉ chuyển sang trạng thái resolved.
   * (draftspect có method này; codebase cũ thiếu.)
   */
  async resolveComment({ anchorText } = {}) {
    if (!this.isAvailable()) return { ok: false };
    return Word.run(async (context) => {
      const results = context.document.body.search(anchorText, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) return { ok: false, error: "Anchor not found" };
      const rangeComments = results.items[0].comments;
      rangeComments.load("items");
      await context.sync();
      if (rangeComments.items.length === 0) return { ok: false, error: "No comment" };
      const c = rangeComments.items[0];
      // Word.js 1.4+: Comment có thuộc tính resolved
      try {
        c.load("resolved");
        await context.sync();
        c.resolved = true;
        await context.sync();
        return { ok: true };
      } catch (err) {
        // Fallback: nếu API resolved không có sẵn, thì xoá comment (giống acceptComment)
        return { ok: false, error: "resolved API không khả dụng: " + err.message };
      }
    });
  }

  /**
   * Đọc replies của 1 comment (threaded comments).
   * (draftspect có listComments + reply riêng; codebase cũ chưa có.)
   */
  async getCommentReplies({ anchorText } = {}) {
    if (!this.isAvailable()) return { ok: false, replies: [] };
    return Word.run(async (context) => {
      const results = context.document.body.search(anchorText, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) return { ok: false, replies: [] };
      const rangeComments = results.items[0].comments;
      rangeComments.load("items/replies");
      await context.sync();
      if (rangeComments.items.length === 0) return { ok: true, replies: [] };
      const replies = rangeComments.items[0].replies;
      replies.load(["content", "author"]);
      await context.sync();
      return {
        ok: true,
        replies: replies.items.map((r) => ({
          author: r.author || "",
          content: r.content || "",
        })),
      };
    });
  }

  /**
   * Bắn hàng loạt nhận xét vào mép lề Word.
   */
  async batchInsertComments(comments) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");
    if (!this.commentsSupported || typeof Word?.Comment === "undefined") {
      throw new Error("Phiên bản Word hiện tại không hỗ trợ chèn bình luận lề.");
    }
    if (!Array.isArray(comments) || comments.length === 0) {
      return { total: 0, success: 0, failed: 0, results: [] };
    }
    return Word.run(async (context) => {
      let success = 0;
      let failed = 0;
      const results = [];
      for (const item of comments) {
        try {
          let targetRange = null;
          if (item.anchor && typeof item.anchor === "string" && item.anchor.trim()) {
            const clean = item.anchor.trim();
            let searchResults = context.document.body.search(clean, { matchCase: false });
            searchResults.load("items");
            await context.sync();
            if (searchResults.items.length > 0) {
              targetRange = searchResults.items[0];
            } else if (clean.length > 25) {
              const prefix = clean.slice(0, 25);
              searchResults = context.document.body.search(prefix, { matchCase: false });
              searchResults.load("items");
              await context.sync();
              if (searchResults.items.length > 0) targetRange = searchResults.items[0];
            }
          }
          if (!targetRange) targetRange = context.document.getSelection();
          targetRange.insertComment(item.text);
          await context.sync();
          success++;
          results.push({ ok: true, anchor: item.anchor });
        } catch (err) {
          failed++;
          results.push({ ok: false, anchor: item.anchor, error: err.message });
        }
      }
      return { total: comments.length, success, failed, results };
    });
  }

  // ─── INSERT (3-tier cascade) ───────────────────────────────────────────

  /**
   * Thay thế vùng chọn.
   * Tier 1: setSelectedDataAsync(HTML) nếu có markdown
   * Tier 2: setSelectedDataAsync(Text)
   * Tier 3: Word.run + insertText (last resort)
   */
  async replaceSelection(text) {
    if (!this.isAvailable()) return;
    const hasMarkdown = /(\*\*|\*|#{1,3}\s|[-*]\s)/.test(text);

    if (hasMarkdown) {
      const html = markdownToHtml(text);
      try {
        await new Promise((resolve, reject) => {
          Office.context.document.setSelectedDataAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
              else reject(new Error(result.error?.message || "HTML insert failed"));
            }
          );
        });
        return;
      } catch (_) {
        /* fall through to text */
      }
    }

    try {
      await new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          text,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message || "Text insert failed"));
          }
        );
      });
      return;
    } catch (_) {
      /* fall through to Word.run */
    }

    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  /**
   * Chèn tại vị trí con trỏ.
   */
  async insertAtCursor(text) {
    if (!this.isAvailable()) return;
    const hasMarkdown = /(\*\*|\*|#{1,3}\s|[-*]\s)/.test(text);

    if (hasMarkdown) {
      const html = markdownToHtml(text);
      try {
        await new Promise((resolve, reject) => {
          Office.context.document.setSelectedDataAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
              else reject(new Error(result.error?.message || "HTML insert failed"));
            }
          );
        });
        return;
      } catch (_) {}
    }

    try {
      await new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          text,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message || "Text insert failed"));
          }
        );
      });
      return;
    } catch (_) {}

    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText(text, Word.InsertLocation.end);
      await context.sync();
    });
  }

  /**
   * Chèn đoạn mới ngay sau đoạn chọn.
   */
  async insertParagraphAfter(text) {
    if (!this.isAvailable()) return;
    const hasMarkdown = /(\*\*|\*|#{1,3}\s|[-*]\s)/.test(text);
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      if (hasMarkdown) {
        try {
          sel.insertHtml(markdownToHtml(text), Word.InsertLocation.after);
          await context.sync();
          return;
        } catch (_) {}
      }
      sel.insertParagraph(text, Word.InsertLocation.after);
      await context.sync();
    });
  }

  /**
   * Tìm anchor trong tài liệu và chèn text ngay phía sau anchor.
   * Hỗ trợ search-anchor contract (matchIndex + searchOptions) từ searchBody().
   *
   * @param {string} anchor
   * @param {string} text
   * @param {Object} options - { asParagraph, style, matchIndex, searchOptions }
   */
  async insertAfterAnchor(anchor, text, { asParagraph = false, style = null, matchIndex = 0, searchOptions = null } = {}) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");
    if (!anchor || !text) throw new Error("Thiếu anchor hoặc text");

    const opts = searchOptions || { matchCase: false, matchWholeWord: false };

    return Word.run(async (context) => {
      const results = context.document.body.search(anchor, opts);
      results.load("items");
      await context.sync();
      if (results.items.length === 0) {
        throw new Error(`Không tìm thấy anchor: "${anchor}"`);
      }
      const idx = Math.max(0, Math.min(matchIndex, results.items.length - 1));
      const targetRange = results.items[idx];
      const parentPara = targetRange.paragraphs.getFirst();
      parentPara.load("text");
      await context.sync();
      const newPara = parentPara.insertParagraph(text, Word.InsertLocation.after);
      await context.sync();
      if (asParagraph && style) {
        newPara.style = style;
        await context.sync();
      }
      return {
        ok: true,
        matchedAt: idx,
        totalMatches: results.items.length,
        anchor_search_text: anchor,
        anchor_match_index: idx,
        anchor_search_options: opts,
      };
    });
  }

  /**
   * Chèn vào cuối tài liệu.
   */
  async appendDocumentEnd(text) {
    if (!this.isAvailable()) return;
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  }

  // ─── TRACK CHANGES (pattern claude-word-addin: toggle flag + insert) ──

  /**
   * Chèn có theo dõi sửa đổi.
   * Pattern đơn giản từ claude-word-addin:
   *   1. Lưu changeTrackingMode hiện tại
   *   2. Bật trackAll
   *   3. Insert text (Word tự sinh redline marks)
   *   4. Khôi phục mode cũ
   *
   * Word tự lo phần <w:ins>/<w:del> — KHÔNG cần build OOXML thủ công.
   */
  async insertWithTrackChanges(oldText, newText, author = "Ori AI") {
    if (!this.isAvailable()) return { ok: false, error: "Office.js không khả dụng" };

    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const currentSelection = sel.text || "";

      // Lưu mode cũ
      let previousMode = null;
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        previousMode = context.document.changeTrackingMode;
      } catch (_) {
        // Mac Word.js: changeTrackingMode không đọc được, bỏ qua
      }

      // Bật track changes
      try {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
      } catch (err) {
        // Mac Word: bật bằng JS API không được, user phải tự bật qua Review ribbon.
        // Fallback: vẫn insert text thường.
        console.warn(
          "[Ori AI] setTrackChanges không khả dụng trên Word này. Hãy bật Track Changes qua Review ribbon trước."
        );
      }

      try {
        if (currentSelection && currentSelection.trim() && newText) {
          // Có selection → replace bằng OOXML fragment chuẩn (không phải full package)
          // Format fragment: <w:p>...</w:p> không cần <w:body> wrapper
          const escapedNew = escapeXml(newText);
          const fragment = `<w:p ${W_NS}><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t xml:space="preserve">${escapedNew}</w:t></w:r></w:p>`;
          sel.insertOoxml(fragment, Word.InsertLocation.replace);
        } else if (newText) {
          // Không có selection → chèn text mới (Word sẽ mark là insertion)
          const escapedNew = escapeXml(newText);
          const fragment = `<w:p ${W_NS}><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t xml:space="preserve">${escapedNew}</w:t></w:r></w:p>`;
          sel.insertOoxml(fragment, Word.InsertLocation.end);
        }
        await context.sync();
      } finally {
        // Luôn khôi phục mode cũ
        if (previousMode !== null) {
          try {
            context.document.changeTrackingMode = previousMode;
            await context.sync();
          } catch (_) {}
        }
      }

      return { ok: true, mode: "track_changes_inserted" };
    });
  }

  // ─── FIND / REPLACE / DELETE ──────────────────────────────────────────

  async findReplace(find, replace, matchCase = false, matchWholeWord = false, maxReplacements = 0) {
    if (!this.isAvailable()) return { matched: 0, replaced: 0 };
    return Word.run(async (context) => {
      const results = context.document.body.search(find, { matchCase, matchWholeWord });
      results.load("items");
      await context.sync();
      const n = maxReplacements > 0 ? Math.min(results.items.length, maxReplacements) : results.items.length;
      for (let i = 0; i < n; i++) {
        results.items[i].insertText(replace, Word.InsertLocation.replace);
      }
      await context.sync();
      return { matched: results.items.length, replaced: n };
    });
  }

  async deleteText(find, maxDeletions = 0) {
    if (!this.isAvailable()) return { matched: 0, deleted: 0 };
    return Word.run(async (context) => {
      const results = context.document.body.search(find, { matchCase: true });
      results.load("items");
      await context.sync();
      const n = maxDeletions > 0 ? Math.min(results.items.length, maxDeletions) : results.items.length;
      for (let i = 0; i < n; i++) {
        results.items[i].delete();
      }
      await context.sync();
      return { matched: results.items.length, deleted: n };
    });
  }

  async setParagraphStyle(anchor, style) {
    if (!this.isAvailable()) return { ok: false };
    return Word.run(async (context) => {
      const results = context.document.body.search(anchor, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) throw new Error(`Không tìm thấy chuỗi neo: ${anchor}`);
      const p = results.items[0].paragraphs.getFirst();
      p.style = style;
      await context.sync();
      return { ok: true };
    });
  }

  /**
   * Chèn OOXML tùy biến tại anchor.
   * OOXML phải là fragment (<w:p>...</w:p>), KHÔNG phải full document package.
   */
  async insertOoxml(anchor, ooxml, location = "after") {
    if (!this.isAvailable()) return { ok: false };
    return Word.run(async (context) => {
      const results = context.document.body.search(anchor, { matchCase: false });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) throw new Error(`Không tìm thấy chuỗi neo: ${anchor}`);
      const locMap = {
        after: Word.InsertLocation.after,
        before: Word.InsertLocation.before,
        replace: Word.InsertLocation.replace,
      };
      const loc = locMap[location] || Word.InsertLocation.after;
      results.items[0].insertOoxml(ooxml, loc);
      await context.sync();
      return { ok: true };
    });
  }

  // ─── TRACKED CHANGES (với timeout bảo vệ Mac Word.js) ────────────────

  /**
   * Liệt kê các sửa đổi Track Changes đang chờ.
   * Có timeout 3s vì getTrackedChanges hay treo trên Mac Word (office-js #5535, #6514).
   */
  async getTrackedChanges(timeoutMs = GET_TRACKED_CHANGES_TIMEOUT_MS) {
    if (!this.isAvailable()) return { count: 0, changes: [], note: "Office.js không khả dụng" };

    const inner = Word.run(async (context) => {
      const changes = context.document.body.getTrackedChanges();
      changes.load(["type", "author", "date", "text"]);
      await context.sync();
      return {
        count: changes.items.length,
        changes: changes.items.map((c, i) => ({
          index: i,
          type: c.type,
          author: c.author,
          date: c.date,
          text: c.text,
        })),
      };
    });

    const result = await withTimeout(inner, timeoutMs, "getTrackedChanges");
    if (!result.ok) {
      console.error("[Ori AI][getTrackedChanges] failed:", result.error);
      return {
        count: 0,
        changes: [],
        note: result.timedOut
          ? `getTrackedChanges timed out sau ${timeoutMs}ms (lỗi Mac Word.js đã biết)`
          : `Lỗi: ${result.error.message}`,
      };
    }
    return result.value;
  }

  /**
   * Bật/tắt Track Changes. Có timeout 5s.
   */
  async setTrackChanges(on) {
    if (!this.isAvailable()) return { on, manual: true };

    const inner = Word.run(async (context) => {
      context.document.changeTrackingMode = on
        ? Word.ChangeTrackingMode.trackAll
        : Word.ChangeTrackingMode.off;
      await context.sync();
      return { on, success: true };
    });

    const result = await withTimeout(inner, SET_TRACK_CHANGES_TIMEOUT_MS, "setTrackChanges");
    if (!result.ok) {
      console.error("[Ori AI][setTrackChanges] failed:", result.error);
      return {
        on,
        manual: true,
        note: "Không thể bật/tắt Track Changes qua API. Vui lòng dùng thanh Review của Word.",
      };
    }
    return result.value;
  }

  /**
   * Đọc thuộc tính tài liệu. Có timeout 4s.
   */
  async getDocumentMetadata() {
    if (!this.isAvailable()) return {};
    const inner = Word.run(async (context) => {
      const props = context.document.properties;
      props.load("title,author,subject,lastAuthor");
      await context.sync();
      return {
        title: props.title || "",
        author: props.author || "",
        subject: props.subject || "",
        lastAuthor: props.lastAuthor || "",
      };
    });
    const result = await withTimeout(inner, GET_DOC_METADATA_TIMEOUT_MS, "getDocumentMetadata");
    if (!result.ok) {
      console.error("[Ori AI][getDocumentMetadata] failed:", result.error);
      return {};
    }
    return result.value;
  }

  /**
   * Lấy "đường dẫn" file Word hiện tại (dùng làm session key trong session-store).
   * Trả về title nếu không lấy được URL/path thật (fallback an toàn).
   * Luôn trả về string, KHÔNG throw.
   */
  async getDocumentPathSafe() {
    if (!this.isAvailable()) return "unknown";
    try {
      const meta = await this.getDocumentMetadata();
      return meta?.title || "untitled";
    } catch (_) {
      return "unknown";
    }
  }

  /**
   * Thống kê số từ, số ký tự và thời gian đọc ước tính.
   */
  getTextAnalytics(text) {
    if (!text || typeof text !== "string") {
      return { words: 0, chars: 0, readTimeMin: 0 };
    }
    const words = text.trim().split(/\s+/).filter(Boolean).length;
    const chars = text.length;
    const readTimeMin = Math.ceil(words / 200);
    return { words, chars, readTimeMin };
  }

  // ─── DECLARATIVE TOOL DISPATCH (word-tool-configs pattern) ───────────

  /**
   * Chạy 1 tool theo tên (dùng word-tool-configs/ làm nguồn truth).
   *
   * Đây là cầu nối giữa:
   *  - MCP server (server/tools.js) — sinh schema từ configs
   *  - UI buttons — có thể gọi thống nhất qua 1 entry
   *  - E2E test — gọi trong Word thật qua config.execute(context, args)
   *
   * @param {string} toolName - tên tool (vd: "word_insert_markdown")
   * @param {Object} args - tham số theo JSON Schema trong config.params
   * @returns {Promise<any>} - kết quả từ tool.execute()
   */
  async executeTool(toolName, args = {}) {
    const { executeWordTool, getWordToolConfig } = await import(
      "./word-tool-configs/index.js?v=2.4.0"
    );
    if (!getWordToolConfig(toolName)) {
      throw new Error(`Tool không tồn tại: ${toolName}`);
    }
    return await executeWordTool(toolName, args);
  }

  // ─── SEARCH-ANCHOR CONTRACT (pi-for-word pattern) ─────────────────────

  /**
   * Tìm text trong body, trả về search-anchor contract để các tool khác dùng.
   *
   * Echo snake_case details (anchor_search_text, anchor_match_index, anchor_search_options)
   * để insertAfterAnchor / insertMarkdown nhận đúng cùng 3 field và chèn chính xác.
   *
   * @param {string} searchText
   * @param {Object} options { matchCase, matchWholeWord, matchIndex, textLimit }
   * @returns {Promise<{ matchCount, matches, anchor_search_text, anchor_match_index, anchor_search_options }>}
   */
  async searchBody(searchText, { matchCase = false, matchWholeWord = false, matchIndex = 0, textLimit = 240 } = {}) {
    if (!this.isAvailable()) return { matchCount: 0, matches: [], anchor_search_text: searchText, anchor_match_index: 0, anchor_search_options: null };
    if (!searchText) throw new Error("searchText là bắt buộc");

    const searchOptions = { matchCase, matchWholeWord };

    return Word.run(async (context) => {
      const results = context.document.body.search(searchText, searchOptions);
      results.load("items");
      await context.sync();
      const count = results.items.length;
      const idx = Math.max(0, Math.min(matchIndex, count - 1));
      const matches = [];
      for (let i = 0; i < count; i++) {
        matches.push({
          index: i,
          text: results.items[i].text?.slice(0, textLimit) || "",
        });
      }
      return {
        matchCount: count,
        matches,
        anchor_search_text: searchText,
        anchor_match_index: idx,
        anchor_search_options: searchOptions,
      };
    });
  }

  // ─── TRACK CHANGES WRAPPER (draftspect pattern) ────────────────────────

  /**
   * Bật/tắt track changes trong scope của 1 async function.
   * Restore mode cũ khi xong (kể cả khi throw).
   *
   * Pattern draftspect/withTrackChanges.
   *
   * @param {Function} body - async (context) => Promise<T>
   * @param {boolean} on - bật (true) hay tắt (false)
   * @returns {Promise<T>}
   */
  async withTrackChanges(on, body) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");

    return Word.run(async (context) => {
      let previousMode = null;
      let wasEnabled = false;

      // Save current mode
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        previousMode = context.document.changeTrackingMode;
      } catch (_) {
        // Mac Word.js không đọc được — bỏ qua
      }

      // Enable/disable
      try {
        context.document.changeTrackingMode = on
          ? Word.ChangeTrackingMode.trackAll
          : Word.ChangeTrackingMode.off;
        await context.sync();
        wasEnabled = true;
      } catch (err) {
        console.warn(
          "[Ori AI] setTrackChanges không khả dụng — fallback sang chế độ thường."
        );
      }

      try {
        return await body(context);
      } finally {
        // Always restore
        if (wasEnabled && previousMode !== null) {
          try {
            context.document.changeTrackingMode = previousMode;
            await context.sync();
          } catch (_) {}
        }
      }
    });
  }
}

export const wordBridge = new WordBridge();
