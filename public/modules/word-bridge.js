/**
 * Word Bridge Module
 * Toàn bộ logic tương tác trực tiếp với Microsoft Word qua Office.js
 */

/* global Office, Word */

export function escapeXml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function buildPlainTextOoxml(text) {
  // Word.js `selection.insertOoxml(xml, location)` chỉ chấp nhận FRAGMENT,
  // không phải full document package. Trả về chuỗi XML gồm 1+ paragraph
  // với namespace w: ở root, Word sẽ merge vào selection.
  const paragraphs = String(text || "").replace(/\r\n/g, "\n").split("\n");
  const bodyXml = paragraphs
    .map((p) => {
      if (!p) {
        return `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>`;
      }
      return (
        `<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="left"/></w:pPr>` +
        `<w:r><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(p)}</w:t></w:r>` +
        `</w:p>`
      );
    })
    .join("");

  return (
    `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    bodyXml +
    `</w:body>`
  );
}

export function buildTrackedChangeOoxml(oldText, newText, author = "AI Assistant") {
  const timestamp = new Date().toISOString();
  const id1 = Math.floor(Math.random() * 1000000) + 1;
  const id2 = id1 + 1;

  let delXml = "";
  if (oldText && oldText.trim()) {
    delXml = `<w:del w:id="${id1}" w:author="${escapeXml(author)}" w:date="${timestamp}">
      <w:r><w:rPr><w:rStyle w:val="DeletedText"/></w:rPr><w:delText xml:space="preserve">${escapeXml(oldText)}</w:delText></w:r>
    </w:del>`;
  }

  const insXml = `<w:ins w:id="${id2}" w:author="${escapeXml(author)}" w:date="${timestamp}">
    <w:r><w:rPr><w:rStyle w:val="InsertedText"/></w:rPr><w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r>
  </w:ins>`;

  const bodyXml = `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${delXml}${insXml}</w:p>`;

  return `<?xml version="1.0" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>${bodyXml}</w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

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
        this._commentsSupported = Office.context.requirements.isSetSupported("WordApi", "1.4");
      } catch {
        this._commentsSupported = false;
      }
    }
    return this._commentsSupported;
  }

  get trackChangesSupported() {
    if (this._trackChangesSupported === null) {
      try {
        this._trackChangesSupported = Office.context.requirements.isSetSupported("WordApi", "1.4");
      } catch {
        this._trackChangesSupported = false;
      }
    }
    return this._trackChangesSupported;
  }

  /**
   * Đọc đoạn văn bản người dùng đang bôi đen (Selection)
   * Sử dụng Office Common API để tăng độ ổn định trên iPad và Word Web
   */
  async getSelectedText() {
    if (!this.isAvailable()) {
      return this.cachedSelection || "";
    }

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
              // Fallback qua Word.run
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
   * Đọc toàn bộ văn bản của tài liệu Word
   */
  async getFullDocumentText(limit = 0) {
    if (!this.isAvailable()) return "";

    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      const list = [];
      for (let i = 0; i < paragraphs.items.length; i++) {
        const t = paragraphs.items[i].text.trim();
        if (t.length > 0) list.push(t);
      }
      const full = list.join("\n\n");
      return limit > 0 && full.length > limit ? full.slice(0, limit) : full;
    });
  }

  /**
   * Đọc danh sách đoạn văn có cấu trúc
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
   * Thuật toán 3 lớp định vị và bôi đen đoạn văn bản (Click-to-Jump)
   */
  async selectRange(anchor, paragraphIndex = null) {
    if (!this.isAvailable()) return { ok: false, error: "Office.js không khả dụng" };

    return Word.run(async (context) => {
      // 1. Nếu có anchor chuỗi
      if (anchor && typeof anchor === "string" && anchor.trim()) {
        const cleanAnchor = anchor.trim();

        // Lớp 1: Khớp chính xác (Exact match)
        let searchResults = context.document.body.search(cleanAnchor, { matchCase: false });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length > 0) {
          searchResults.items[0].select();
          await context.sync();
          return { ok: true, matched: "exact", count: searchResults.items.length };
        }

        // Lớp 2: Khớp mờ phần đầu (Prefix search 30 ký tự)
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

      // Lớp 3: Định vị theo chỉ mục đoạn văn (Paragraph Index)
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

      throw new Error(`Không tìm thấy vị trí neo trong tài liệu: '${anchor || paragraphIndex}'`);
    });
  }

  /**
   * Chèn bình luận mép lề phải (Margin Comments via WordApi 1.4)
   * Yêu cầu Word 2016+ / Microsoft 365. Không có fallback inline
   * (vì sẽ làm thay đổi nội dung tài liệu gốc — trái với triết lý
   * "không xáo trộn bài viết").
   */
  async insertComment(commentText, anchor = null) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");

    if (!this.commentsSupported || typeof Word?.Comment === "undefined") {
      throw new Error(
        "Phiên bản Word hiện tại không hỗ trợ chèn bình luận (cần Word 2016 trở lên hoặc Microsoft 365)."
      );
    }

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

      if (!targetRange) {
        targetRange = context.document.getSelection();
      }

      targetRange.insertComment(commentText);
      await context.sync();
      return { ok: true, type: "margin_comment" };
    });
  }

  /**
   * Thay thế vùng chọn
   */
  async replaceSelection(text) {
    if (!this.isAvailable()) return;
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertOoxml(buildPlainTextOoxml(text), Word.InsertLocation.replace);
      await context.sync();
    });
  }

  /**
   * Chèn tại vị trí con trỏ
   */
  async insertAtCursor(text) {
    if (!this.isAvailable()) return;
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText(text, Word.InsertLocation.end);
      await context.sync();
    });
  }

  /**
   * Chèn đoạn mới ngay sau đoạn chọn
   */
  async insertParagraphAfter(text) {
    if (!this.isAvailable()) return;
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertParagraph(text, Word.InsertLocation.after);
      await context.sync();
    });
  }

  /**
   * Tìm anchor trong tài liệu và chèn text (hoặc paragraph mới) ngay phía sau anchor.
   * Gom tất cả vào 1 Word.run block để giữ state consistent.
   *
   * @param {string} anchor - Chuỗi cần tìm
   * @param {string} text - Nội dung cần chèn
   * @param {Object} options - { asParagraph: bool, style: string|null }
   */
  async insertAfterAnchor(anchor, text, { asParagraph = false, style = null } = {}) {
    if (!this.isAvailable()) throw new Error("Office.js không khả dụng");
    if (!anchor || !text) throw new Error("Thiếu anchor hoặc text");

    return Word.run(async (context) => {
      const results = context.document.body.search(anchor, { matchCase: false });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        throw new Error(`Không tìm thấy anchor: "${anchor}"`);
      }

      // Lấy paragraph cha của vị trí match (tránh lệch khi anchor nằm giữa paragraph)
      const parentPara = results.items[0].paragraphs.getFirst();
      parentPara.load("text");
      await context.sync();

      // Chèn ngay sau paragraph chứa anchor
      const newPara = parentPara.insertParagraph(text, Word.InsertLocation.after);
      await context.sync();

      // Nếu yêu cầu paragraph mới riêng + style → style cho paragraph mới
      if (asParagraph && style) {
        newPara.style = style;
        await context.sync();
      }

      return { ok: true, matchedAt: 0, totalMatches: results.items.length };
    });
  }

  /**
   * Chèn vào cuối tài liệu
   */
  async appendDocumentEnd(text) {
    if (!this.isAvailable()) return;
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  }

  /**
   * Chèn có theo dõi sửa đổi (Track Changes / Redlines)
   */
  async insertWithTrackChanges(newText, mode = "replace") {
    if (!this.isAvailable()) return;

    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const oldText = sel.text || "";

      // Sử dụng hybrid OOXML để đảm bảo hiện vạch đỏ trên mọi nền tảng
      const ooxml = buildTrackedChangeOoxml(oldText, newText);
      sel.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
      return { ok: true };
    });
  }

  /**
   * Tìm và thay thế chuỗi
   */
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

  /**
   * Xóa văn bản
   */
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

  /**
   * Gán kiểu định dạng đoạn (Heading, Normal, v.v.)
   * Sử dụng matchCase: false để consistent với selectRange, insertComment, insertAfterAnchor.
   */
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
   * Chèn OOXML tùy biến
   * Sử dụng matchCase: false để consistent với các search khác.
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

  /**
   * Liệt kê các sửa đổi Track Changes đang chờ
   */
  async getTrackedChanges(timeoutMs = 4000) {
    if (!this.isAvailable()) return { count: 0, changes: [] };

    return Word.run(async (context) => {
      try {
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
      } catch (_) {
        return { count: 0, changes: [], note: "Môi trường Word hiện tại không hỗ trợ đọc danh sách Track Changes" };
      }
    });
  }

  /**
   * Bật/tắt chế độ Track Changes
   */
  async setTrackChanges(on) {
    if (!this.isAvailable()) return { on };

    return Word.run(async (context) => {
      try {
        context.document.changeTrackingMode = on
          ? Word.ChangeTrackingMode.trackAll
          : Word.ChangeTrackingMode.off;
        await context.sync();
        return { on, success: true };
      } catch (_) {
        return { on, manual: true, note: "Vui lòng bật/tắt Track Changes qua thanh công cụ Review của Word" };
      }
    });
  }

  /**
   * Đọc thuộc tính tài liệu
   */
  async getDocumentMetadata() {
    if (!this.isAvailable()) return {};

    return Word.run(async (context) => {
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
  }

  /**
   * Thống kê số từ, số ký tự và thời gian đọc ước tính
   */
  getTextAnalytics(text) {
    if (!text || typeof text !== "string") {
      return { words: 0, chars: 0, readTimeMin: 0 };
    }
    const words = text.trim().split(/\s+/).filter(Boolean).length;
    const chars = text.length;
    const readTimeMin = Math.ceil(words / 200); // 200 từ/phút
    return { words, chars, readTimeMin };
  }
}

export const wordBridge = new WordBridge();
