# 📝 AI Word Add-in & MCP Agent Bridge

> **Bộ Trợ lý AI Soạn thảo Microsoft Word Toàn diện** — Kết hợp hoàn hảo giữa **Giao diện Sidebar (Taskpane) tương tác trực tiếp cho người dùng** và **Cổng Cầu nối MCP Server (Model Context Protocol) cho các AI Agent bên ngoài** (Claude Code, Cursor, Codex, OpenCode) điều khiển Word trực tiếp.

[![Platform: Office.js](https://img.shields.io/badge/Office.js-WordApi%201.4-blue.svg)](https://docs.microsoft.com/en-us/office/dev/add-ins/)
[![MCP: Supported](https://img.shields.io/badge/MCP-Model%20Context%20Protocol-green.svg)](https://modelcontextprotocol.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 🌟 Điểm Nhấn Độc Đáo

1. **Dual-Mode Song Song (2 Chế độ trong 1):**
   * **Chế độ Người dùng (Taskpane Sidebar):** Giao diện tiếng Việt hiện đại, thanh kỹ năng 1-chạm (Viết lại, Sửa ngữ pháp, Gọt từ thừa, Trang trọng hóa, Tóm tắt, Nhận xét lề trang...).
   * **Chế độ Agent Ngoài (MCP Server):** Cung cấp 14+ tools chuẩn MCP giúp các AI Agent bên ngoài đọc cấu trúc tài liệu, tìm kiếm/thay thế, chèn bình luận, đổi định dạng và điều khiển con trỏ trong Word theo thời gian thực.
2. **🎯 Điều hướng Click-to-Jump:** Khi AI nhận xét hoặc sửa đoạn nào, nhấp vào trích dẫn là Word tự động cuộn màn hình và bôi đen chính xác câu văn đó trong tài liệu.
3. **💬 Bình luận Lề Trang (Margin Comments):** AI để lại nhận xét, góp ý ở mép phải tài liệu như một biên tập viên mà không làm xáo trộn bài viết gốc.
4. **🔴 Theo dõi Sửa đổi (Track Changes / Redlines):** Mọi văn bản AI thay đổi được hiển thị bằng vạch đỏ Redlines trực quan.
5. **🧠 Giọng văn Tham chiếu (House Voice) & Nhớ theo File:** AI học theo văn phong mẫu của người dùng; mỗi file `.docx` tự lưu giữ lịch sử chat và cài đặt riêng.
6. **⚡ Tích hợp AI Endpoint Tùy biến & Streaming SSE:** Kết nối trực tiếp tới AI API Endpoint có sẵn với tốc độ chữ chạy mượt mà từng token và hỗ trợ Thinking Mode (DeepSeek R1 / Claude 3.7).

---

## 🏗️ Kiến Trúc Hệ Thống

```
                  ┌─────────────────────────────────────────────────────────────┐
                  │                 HỆ THỐNG AI WORD ADD-IN                     │
                  └──────────────────────────────┬──────────────────────────────┘
                                                 │
        ┌────────────────────────────────────────┴────────────────────────────────────────┐
        ▼                                                                                 ▼
┌─────────────────────────┐                                                   ┌─────────────────────────┐
│  CHẾ ĐỘ 1: TASKPANE UI  │                                                   │   CHẾ ĐỘ 2: MCP BRIDGE  │
│  (Người dùng trực tiếp  │                                                   │  (Dành cho AI Agent:    │
│   thao tác trong Word)  │                                                   │  Claude Code, Cursor...)│
└───────────┬─────────────┘                                                   └───────────┬─────────────┘
            │                                                                             │
            │ Giao diện Sidebar Office.js                                                 │ REST (/op, /ops) & WebSocket (/ws)
            │ (Click-to-Jump, Margin Comment, Skills, State)                              │ & Model Context Protocol (MCP)
            │                                                                             │
            └────────────────────────────────────┬────────────────────────────────────────┘
                                                 │
                                                 ▼
                              ┌─────────────────────────────────────┐
                              │     AI BACKEND / BRIDGE SERVER      │
                              │     (Node.js / Express / ws)        │
                              └──────────────────┬──────────────────┘
                                                 │
                        ┌────────────────────────┴────────────────────────┐
                        ▼                                                 ▼
         ┌─────────────────────────────┐                   ┌─────────────────────────────┐
         │   MICROSOFT WORD DOCUMENT   │                   │    USER'S AI API ENDPOINT   │
         │   (Office.js Live Runtime)  │                   │   (OpenAI / Anthropic /     │
         │                             │                   │    DeepSeek / Custom URL)   │
         └─────────────────────────────┘                   └─────────────────────────────┘
```

---

## 📁 Cấu Trúc Thư Mục

```
ai-add-in-word/
├── manifest.xml                # Office Add-in manifest chuẩn (dùng để nạp vào Word)
├── package.json                # Cấu hình dự án & dependencies
├── server/                     # BACKEND SERVER, AI PROXY & MCP SERVER
│   ├── index.js                # Server chính (Express + WebSocket + MCP Bridge)
│   ├── ai-proxy.js             # Bộ chuyển tiếp SSE tới Custom AI Endpoint
│   ├── mcp-server.js           # Chuẩn Model Context Protocol Server
│   └── tools-catalog.js        # Khai báo schema 14+ Tools thao tác Word
├── public/                     # GIAO DIỆN TASKPANE SIDEBAR
│   ├── index.html              # HTML giao diện Sidebar tiếng Việt
│   ├── styles.css              # Giao diện Dark/Light mode cao cấp
│   ├── app.js                  # Điều khiển giao diện & tương tác người dùng
│   ├── skills.json             # Danh mục kỹ năng 1-chạm (Viết lại, Sửa lỗi...)
│   └── modules/                # CÁC MODULE OFFICE.JS TINH HOA
│       ├── word-bridge.js      # Click-to-Jump, Margin Comments, Track Changes
│       ├── doc-state.js        # Lưu nhớ lịch sử chat theo từng file .docx
│       ├── sse-stream.js       # Phân giải luồng stream SSE token-by-token
│       └── ws-client.js        # Lắng nghe lệnh từ MCP Server
└── docs/                       # TÀI LIỆU DỰ ÁN
    ├── AI_WORD_ADDIN_FEATURE_SPEC.md     # Bản đặc tả tính năng & Checklist đối chiếu
    ├── ARCHITECTURE_AND_SYSTEM_DESIGN.md # Bản thiết kế kiến trúc & giải pháp kỹ thuật
    ├── USER_GUIDE.md                     # Hướng dẫn sử dụng cho người dùng (Chị gái)
    └── MCP_INTEGRATION_GUIDE.md          # Hướng dẫn kết nối MCP cho AI Agent
```

---

## 🚀 Hướng Dẫn Cài Đặt & Khởi Chạy

### 1. Yêu cầu Môi trường
* **Node.js** 18.0 trở lên.
* **Microsoft Word** (Word 365 Web trên trình duyệt, Word Desktop trên Windows hoặc macOS).

### 2. Cài đặt Dependencies & Khởi chạy Server
```bash
# Cài đặt các gói phụ thuộc
npm install

# Khởi chạy server tích hợp (Add-in Backend + MCP Server)
npm start
```
Server sẽ chạy mặc định tại: `https://localhost:3001` (hoặc `http://localhost:3001`).

---

## 📥 Cách Nạp Add-in Vào Microsoft Word (Sideloading)

### Cách 1: Nạp vào Word 365 Web (Đơn giản nhất, dùng được ngay)
1. Mở trình duyệt và truy cập [word.office.com](https://word.office.com), mở bất kỳ tài liệu nào.
2. Trên thanh menu, chọn thẻ **Insert (Chèn)** ➔ **Add-ins (Tiện ích bổ sung)** ➔ **Upload My Add-in (Tải tiện ích của tôi lên)**.
3. Chọn file `manifest.xml` từ thư mục dự án.
4. Nút **Trợ lý AI** sẽ xuất hiện trên thanh công cụ Home Ribbon. Nhấp vào để mở Sidebar!

### Cách 2: Nạp vào Word Desktop trên Windows
1. Mở Word ➔ Vào **File ➔ Options ➔ Trust Center ➔ Trust Center Settings... ➔ Trusted Add-in Catalogs**.
2. Thêm đường dẫn thư mục chứa file `manifest.xml`, tích chọn **Show in Menu**, bấm **OK**.
3. Khởi động lại Word ➔ Vào **Insert ➔ My Add-ins ➔ Shared Folder** ➔ Chọn **AI Word Assistant**.

### Cách 3: Nạp vào Word Desktop trên macOS
Copy file `manifest.xml` vào thư mục Wef của Word:
```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```
Khởi động lại Word ➔ Vào **Insert ➔ My Add-ins ➔ Developer Add-ins** ➔ Chọn **AI Word Assistant**.

---

## 🤖 Kết Nối AI Agent Ngoài (Claude Code, Cursor, OpenCode qua MCP)

Cấu hình MCP Server vào tệp cấu hình của Agent (ví dụ: `claude_desktop_config.json` hoặc Cursor MCP settings):

```json
{
  "mcpServers": {
    "word-bridge": {
      "command": "node",
      "args": ["/đường/dẫn/tới/ai-add-in-word/server/mcp-server.js"]
    }
  }
}
```

Xem chi tiết danh mục 14+ tools và hướng dẫn sử dụng trong [docs/MCP_INTEGRATION_GUIDE.md](docs/MCP_INTEGRATION_GUIDE.md).

---

## 📚 Tài Liệu Tham Khảo Thêm
* 📖 [Sổ Tay Hướng Dẫn Sử Dụng Cho Người Dùng (USER_GUIDE.md)](docs/USER_GUIDE.md)
* ⚙️ [Tài Liệu Tích Hợp Giao Thức MCP (MCP_INTEGRATION_GUIDE.md)](docs/MCP_INTEGRATION_GUIDE.md)
* 📐 [Thiết Kế Kiến Trúc & Giải Pháp Kỹ Thuật (ARCHITECTURE_AND_SYSTEM_DESIGN.md)](docs/ARCHITECTURE_AND_SYSTEM_DESIGN.md)
* ✅ [Bản Đặc Tả Tính Năng & Checklist Nghiệm Thu (AI_WORD_ADDIN_FEATURE_SPEC.md)](docs/AI_WORD_ADDIN_FEATURE_SPEC.md)

---

## 📄 Bản Quyền
Phát hành theo giấy phép [MIT License](LICENSE).
