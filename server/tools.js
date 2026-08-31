/**
 * Tool Catalog & Dispatcher cho Word Bridge / MCP Server
 * Định nghĩa 14+ công cụ thao tác Word chuẩn JSON Schema và điều phối lệnh qua WebSocket
 */

import { v4 as uuidv4 } from "uuid";

// Bộ quản lý kết nối WebSocket của Word Taskpane
class WordClientManager {
  constructor() {
    this.clients = new Map(); // id -> { ws, kind, info, lastSeen }
    this.pendingOps = new Map(); // opId -> { resolve, reject, timeout }
  }

  register(ws, meta = {}) {
    const id = uuidv4();
    const client = {
      id,
      ws,
      kind: meta.kind || "word",
      info: meta.info || "Word Taskpane",
      connectedAt: new Date().toISOString(),
      lastSeen: Date.now(),
    };
    this.clients.set(id, client);

    ws.on("close", () => {
      this.clients.delete(id);
    });

    return id;
  }

  getActiveClient(targetId) {
    if (targetId && this.clients.has(targetId)) {
      return this.clients.get(targetId);
    }
    // Lấy client đang mở đầu tiên
    for (const client of this.clients.values()) {
      if (client.ws.readyState === 1) {
        return client;
      }
    }
    return null;
  }

  handleResult(message) {
    const { id, ok, result, error, detail } = message;
    if (this.pendingOps.has(id)) {
      const { resolve, reject, timeout } = this.pendingOps.get(id);
      clearTimeout(timeout);
      this.pendingOps.delete(id);

      if (ok) {
        resolve(result);
      } else {
        const err = new Error(error || "Thao tác Word thất bại");
        err.detail = detail;
        reject(err);
      }
    }
  }

  async dispatchOp(op, { targetClientId, timeoutMs = 15000 } = {}) {
    const client = this.getActiveClient(targetClientId);
    if (!client) {
      throw new Error("Không có Microsoft Word Add-in nào đang kết nối. Vui lòng mở Add-in trong Word trước.");
    }

    const id = uuidv4();
    const payload = {
      type: "op",
      id,
      op,
    };

    return new Promise((resolve, reject) => {
      // Validate ws state TRƯỚC khi register pending (tránh leak pending nếu ws đã đóng)
      if (client.ws.readyState !== 1) {
        return reject(new Error("WebSocket client không sẵn sàng (đã đóng hoặc đang kết nối)."));
      }

      // Tạo timeout
      const timeout = setTimeout(() => {
        if (this.pendingOps.has(id)) {
          this.pendingOps.delete(id);
          reject(new Error(`Thao tác '${op.kind}' quá thời gian chờ (${timeoutMs}ms)`));
        }
      }, timeoutMs);

      // Set pending TRƯỚC khi send để đảm bảo handleResult tìm thấy
      this.pendingOps.set(id, { resolve, reject, timeout });

      // Gửi op
      try {
        client.ws.send(JSON.stringify(payload), (err) => {
          // Nếu send fail (vd: ws đã đóng giữa lúc check và send),
          // cleanup pending để tránh leak.
          if (err) {
            if (this.pendingOps.has(id)) {
              this.pendingOps.delete(id);
              clearTimeout(timeout);
              reject(err);
            }
          }
        });
      } catch (err) {
        // Send throw đồng bộ (hiếm gặp nhưng cần handle)
        if (this.pendingOps.has(id)) {
          this.pendingOps.delete(id);
          clearTimeout(timeout);
        }
        reject(err);
      }
    });
  }

  getStatus() {
    const list = [];
    for (const [id, c] of this.clients.entries()) {
      list.push({
        id,
        kind: c.kind,
        info: c.info,
        connectedAt: c.connectedAt,
        ready: c.ws.readyState === 1,
      });
    }
    return {
      connectedCount: list.length,
      clients: list,
    };
  }
}

export const clientManager = new WordClientManager();

// Danh mục 14+ Tools chuẩn MCP & REST
export const tools = [
  {
    name: "word_ping",
    description: "Kiểm tra kết nối 2 chiều giữa AI Agent và Microsoft Word Taskpane.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "ping" },
      },
    },
  },
  {
    name: "word_getText",
    description: "Đọc toàn bộ văn bản thô của tài liệu Word đang mở.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "getText" },
        limit: { type: "integer", default: 8000, description: "Giới hạn số ký tự đọc (mặc định 8000)." },
      },
    },
  },
  {
    name: "word_getParagraphs",
    description: "Đọc danh sách các đoạn văn có cấu trúc trong tài liệu Word kèm định dạng kiểu (style), text tóm tắt và độ dài.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "getParagraphs" },
        styleFilter: { type: "string", description: "Lọc theo tên kiểu (ví dụ: 'Heading 1', 'Normal')." },
        limit: { type: "integer", default: 0, description: "Số đoạn tối đa cần đọc (0 = tất cả)." },
        textLimit: { type: "integer", default: 300, description: "Độ dài ký tự tối đa cho mỗi đoạn trích." },
      },
    },
  },
  {
    name: "word_getSelection",
    description: "Đọc đoạn văn bản người dùng đang bôi đen trên màn hình Word.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "getSelection" },
      },
    },
  },
  {
    name: "word_selectRange",
    description: "Điều khiển Word tự động cuộn màn hình và bôi đen (highlight) đoạn văn bản theo từ khóa 'anchor' hoặc chỉ mục đoạn 'paragraphIndex'.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "selectRange" },
        anchor: { type: "string", description: "Từ khóa hoặc đoạn câu cần định vị và bôi đen." },
        paragraphIndex: { type: "integer", description: "Vị trí chỉ mục đoạn văn (nếu không dùng anchor)." },
      },
    },
  },
  {
    name: "word_findReplace",
    description: "Tìm kiếm và thay thế chuỗi trên toàn bộ tài liệu Word (có thể ghi nhận Track Changes).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "findReplace" },
        find: { type: "string", description: "Chuỗi cần tìm." },
        replace: { type: "string", description: "Chuỗi thay thế." },
        matchCase: { type: "boolean", default: false, description: "Phân biệt hoa thường." },
        matchWholeWord: { type: "boolean", default: false, description: "Khớp toàn bộ từ." },
        maxReplacements: { type: "integer", default: 0, description: "Số lượng thay thế tối đa (0 = tất cả)." },
      },
      required: ["find", "replace"],
    },
  },
  {
    name: "word_insertAfterText",
    description: "Tìm một vị trí neo ('anchor') trong văn bản và chèn nội dung mới ngay phía sau.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "insertAfterText" },
        anchor: { type: "string", description: "Vị trí chuỗi neo." },
        text: { type: "string", description: "Văn bản mới cần chèn." },
        asParagraph: { type: "boolean", default: false, description: "Tạo thành một đoạn văn mới." },
        style: { type: "string", description: "Kiểu định dạng đoạn (chỉ dùng khi asParagraph=true)." },
      },
      required: ["anchor", "text"],
    },
  },
  {
    name: "word_deleteText",
    description: "Tìm và xóa đoạn văn bản chỉ định khỏi tài liệu Word.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "deleteText" },
        find: { type: "string", description: "Chuỗi cần xóa." },
        maxDeletions: { type: "integer", default: 0, description: "Số lần xóa tối đa (0 = tất cả)." },
      },
      required: ["find"],
    },
  },
  {
    name: "word_setParagraphStyle",
    description: "Gán định dạng kiểu cho đoạn văn chứa chuỗi neo (Heading 1, Heading 2, Quote, Normal...).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "setParagraphStyle" },
        anchor: { type: "string", description: "Chuỗi định vị đoạn văn." },
        style: { type: "string", description: "Tên kiểu định dạng (ví dụ: 'Heading 1', 'Normal')." },
      },
      required: ["anchor", "style"],
    },
  },
  {
    name: "word_insertComment",
    description: "Tạo một bình luận (Word Comment) ở mép lề phải tài liệu tại đoạn văn bản chỉ định.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "insertComment" },
        anchor: { type: "string", description: "Đoạn văn bản neo cần nhận xét." },
        commentText: { type: "string", description: "Nội dung lời nhận xét / bình luận." },
      },
      required: ["commentText"],
    },
  },
  {
    name: "word_insertOoxml",
    description: "Chèn đoạn mã OpenXML (OOXML) tùy biến vào vị trí neo (dùng cho bảng biểu, công thức hoặc thẻ <w:ins>).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "insertOoxml" },
        anchor: { type: "string", description: "Vị trí chuỗi neo." },
        ooxml: { type: "string", description: "Chuỗi mã OOXML XML hợp lệ." },
        location: { type: "string", enum: ["before", "after", "replace"], default: "after" },
      },
      required: ["anchor", "ooxml"],
    },
  },
  {
    name: "word_getTrackedChanges",
    description: "Liệt kê tất cả các sửa đổi đang chờ duyệt trong tài liệu Word (tác giả, ngày giờ, loại sửa đổi, văn bản).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "getTrackedChanges" },
        timeoutMs: { type: "integer", default: 4000 },
      },
    },
  },
  {
    name: "word_setTrackChanges",
    description: "Bật hoặc tắt chế độ ghi nhận sửa đổi (Track Changes) trong Word.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "setTrackChanges" },
        on: { type: "boolean", description: "true để bật, false để tắt." },
      },
      required: ["on"],
    },
  },
  {
    name: "word_getMetadata",
    description: "Đọc thông tin thuộc tính của tài liệu Word (Title, Author, Subject, Last Author).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", const: "getMetadata" },
      },
    },
  },
];
