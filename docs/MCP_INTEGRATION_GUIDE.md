# 🤖 MCP INTEGRATION GUIDE
## DỰ ÁN: MODEL CONTEXT PROTOCOL BRIDGE FOR MICROSOFT WORD

> **Dành cho Lập trình viên & AI Agents:** Tài liệu đặc tả kỹ thuật cách kết nối các AI Agent bên ngoài (như **Claude Code, Cursor, Codex, OpenCode, Cline, AutoGen**) với tài liệu Word đang mở thông qua chuẩn **Model Context Protocol (MCP)**.

---

# MỤC LỤC
1. [TỔNG QUAN KIẾN TRÚC MCP BRIDGE](#1-tổng-quan-kiến-trúc-mcp-bridge)
2. [CẤU HÌNH KẾT NỐI MCP CLIENT](#2-cấu-hình-kết-nối-mcp-client)
3. [DANH MỤC CHI TIẾT 14+ TOOLS CHUẨN MCP](#3-danh-mục-chi-tiết-14-tools-chuẩn-mcp)
4. [HƯỚNG DẪN GỌI LỆNH & VÍ DỤ MẪU (EXAMPLES)](#4-hướng-dẫn-gọi-lệnh--ví-dụ-mẫu-examples)
5. [CƠ CHẾ BATCH OPERATIONS & XỬ LÝ LỖI](#5-cơ-chế-batch-operations--xử-lý-lỗi)

---

# 1. TỔNG QUAN KIẾN TRÚC MCP BRIDGE

```
┌────────────────────────────────────────┐
│           EXTERNAL AI AGENT            │  (Claude Code / Cursor / Codex / MCP Client)
└──────────────────┬─────────────────────┘
                   │ MCP Transport (JSON-RPC over stdio hoặc SSE)
                   ▼
┌────────────────────────────────────────┐
│           LOCAL MCP SERVER             │  (server/mcp-server.js)
└──────────────────┬─────────────────────┘
                   │ WebSocket Push (ws://localhost:3001/ws)
                   ▼
┌────────────────────────────────────────┐
│        WORD TASKPANE OFFICE.JS         │  (public/modules/ws-client.js)
└──────────────────┬─────────────────────┘
                   │ Live Office.js API
                   ▼
┌────────────────────────────────────────┐
│        ACTIVE WORD DOCUMENT (.docx)    │
└────────────────────────────────────────┘
```

---

# 2. CẤU HÌNH KẾT NỐI MCP CLIENT

### Cấu hình cho Claude Desktop (`claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "word-assistant": {
      "command": "node",
      "args": ["/đường/dẫn/tuyệt/đối/tới/ai-add-in-word/server/mcp-server.js"]
    }
  }
}
```

### Cấu hình cho Cursor IDE (`.cursor/mcp.json`):
```json
{
  "mcpServers": {
    "word-bridge": {
      "command": "node",
      "args": ["server/mcp-server.js"]
    }
  }
}
```

---

# 3. DANH MỤC CHI TIẾT 14+ TOOLS CHUẨN MCP

| Tên Tool | Host | Mô tả chức năng |
| :--- | :--- | :--- |
| `word_ping` | `any` | Kiểm tra kết nối 2 chiều giữa Agent và Word Taskpane. |
| `word_getText` | `word` | Đọc toàn bộ văn bản thô trong tài liệu (hỗ trợ tham số `limit`). |
| `word_getParagraphs` | `word` | Đọc danh sách đoạn văn có cấu trúc kèm kiểu định dạng (Heading, Normal...). |
| `word_getSelection` | `word` | Đọc đoạn văn bản người dùng đang bôi đen trên màn hình Word. |
| `word_selectRange` | `word` | **Điều khiển Word tự cuộn và bôi đen đoạn văn bản theo từ khóa `anchor` hoặc `paragraphIndex`.** |
| `word_findReplace` | `word` | Tìm kiếm và thay thế chuỗi (chính xác/hoa thường/toàn bộ từ/giới hạn số lượng). |
| `word_insertAfterText` | `word` | Tìm vị trí neo (`anchor`) và chèn văn bản/đoạn văn mới phía sau. |
| `word_deleteText` | `word` | Tìm và xóa đoạn văn bản chỉ định. |
| `word_setParagraphStyle` | `word` | Đổi định dạng kiểu đoạn văn (Heading 1, Heading 2, Quote, Normal). |
| `word_insertComment` | `word` | **Tạo bình luận lề phải Word tại đoạn văn bản chỉ định.** |
| `word_insertOoxml` | `word` | Chèn mã OOXML tùy biến (bảng biểu phức tạp, định dạng đặc biệt, thẻ `<w:ins>`). |
| `word_getTrackedChanges` | `word` | Liệt kê tất cả các sửa đổi đang chờ duyệt (tác giả, ngày giờ, loại sửa đổi, nội dung). |
| `word_setTrackChanges` | `word` | Bật/tắt chế độ ghi nhận sửa đổi trong Word. |
| `word_getMetadata` | `word` | Đọc thông tin thuộc tính file (Title, Author, Subject, Last Author). |

---

# 4. HƯỚNG DẪN GỌI LỆNH & VÍ DỤ MẪU (EXAMPLES)

### Ví dụ 1: Đọc cấu trúc các đoạn văn bản trong tài liệu
**Yêu cầu Agent:**
```json
{
  "name": "word_getParagraphs",
  "arguments": {
    "styleFilter": "Heading 1",
    "textLimit": 200
  }
}
```
**Phản hồi từ Word:**
```json
{
  "total": 45,
  "returned": 3,
  "paragraphs": [
    { "index": 0, "style": "Heading 1", "text": "1. Giới thiệu tổng quan", "length": 23 },
    { "index": 12, "style": "Heading 1", "text": "2. Phương pháp thực hiện", "length": 25 },
    { "index": 30, "style": "Heading 1", "text": "3. Kết luận & Kiến nghị", "length": 24 }
  ]
}
```

---

### Ví dụ 2: Điều khiển Word nhảy tới và bôi đen một đoạn văn
**Yêu cầu Agent:**
```json
{
  "name": "word_selectRange",
  "arguments": {
    "anchor": "Phương pháp thực hiện"
  }
}
```
**Kết quả:** Màn hình Word tự động cuộn đến trang chứa tiêu đề "2. Phương pháp thực hiện" và bôi đen đoạn văn đó.

---

### Ví dụ 3: Thêm bình luận lề phải trang (Margin Comment)
**Yêu cầu Agent:**
```json
{
  "name": "word_insertComment",
  "arguments": {
    "anchor": "Phương pháp thực hiện",
    "commentText": "Cần bổ sung thêm biểu đồ so sánh dữ liệu thực nghiệm ở mục này."
  }
}
```
**Kết quả:** Một Word Comment xuất hiện ngay bên lề phải của đoạn văn được chỉ định.

---

### Ví dụ 4: Tìm và thay thế chuỗi có ghi nhận Track Changes
**Yêu cầu Agent:**
```json
{
  "name": "word_findReplace",
  "arguments": {
    "find": "kế hoạch 2025",
    "replace": "kế hoạch 2026",
    "matchCase": true
  }
}
```

---

# 5. CƠ CHẾ BATCH OPERATIONS & XỬ LÝ LỖI

### Gọi nhiều thao tác liên tiếp (`/ops` Endpoint):
Agent có thể gửi một mảng các thao tác để thực thi theo thứ tự nguyên tử:
```json
{
  "ops": [
    { "kind": "findReplace", "find": "Dự thảo", "replace": "Chính thức" },
    { "kind": "setParagraphStyle", "anchor": "Chính thức", "style": "Heading 1" },
    { "kind": "insertComment", "anchor": "Chính thức", "commentText": "Đã phê duyệt bản chính thức." }
  ],
  "stopOnError": true
}
```

### Báo cáo Lỗi Tự Sửa Sai (Self-Correction Schema):
Nếu từ khóa `anchor` không tồn tại trong tài liệu, hệ thống trả về thông báo lỗi chuẩn xác:
```json
{
  "ok": false,
  "error": "anchor not found: 'Phương thức triển khai'",
  "detail": {
    "suggestion": "Thử tìm kiếm với đoạn trích ngắn hơn hoặc sử dụng word_getParagraphs để kiểm tra danh sách văn bản hiện có."
  }
}
```
Agent có thể dựa vào thông báo này để tự động điều chỉnh câu lệnh và thử lại.
