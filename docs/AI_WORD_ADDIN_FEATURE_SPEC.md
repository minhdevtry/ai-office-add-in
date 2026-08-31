# BẢN ĐẶC TẢ TÍNH NĂNG & CHECKLIST HỆ THỐNG
## DỰ ÁN: AI WORD ADD-IN & MCP AGENT BRIDGE

> **Tài liệu này được xây dựng dựa trên việc chắt lọc những tinh hoa, tính năng mạnh mẽ và kiến trúc sạch nhất từ các dự án tham khảo cốt lõi trong `.ref`:**
> 1. `claude-word-addin`: Giao diện Taskpane, Track Changes (Redlines), Streaming token-by-token.
> 2. `vivword`: Bình luận lề trang (Margin Comments), Hệ thống Skills 1-chạm, Giọng văn tham chiếu (House Voice), Lưu nhớ theo file Word (`document.settings`).
> 3. `office-copilot-lite`: Kết nối AI Gateway đa năng (kết nối endpoint tùy biến của người dùng).
> 4. `Claude-Word-Plugin`: Thẻ đính kèm ngữ cảnh (Context Attachment Chips), Chế độ Suy nghĩ sâu (Extended Thinking).
> 5. `wordbridge`: Cổng Headless Agent Bridge, WebSocket & REST API, Bộ công cụ chuẩn MCP cho Word.

---

# MỤC LỤC
1. [TỔNG HỢP VÀ ĐỐI CHIẾU CÁC REPO CỐT LÕI](#1-tổng-hợp-và-đối-chiếu-các-repo-cốt-lõi)
2. [KIẾN TRÚC TỔNG THỂ HỆ THỐNG](#2-kiến-trúc-tổng-thể-hệ-thống)
3. [DANH MỤC TÍNH NĂNG CHI TIẾT (FEATURE SPECIFICATION)](#3-danh-mục-tính-năng-chi-tiết-feature-specification)
   - [Trụ cột 1: Tương tác Word Chuyên sâu & Điều hướng (Word-Native & Navigation)](#trụ-cột-1-tương-tác-word-chuyên-sâu--điều-hướng-word-native--navigation)
   - [Trụ cột 2: Engine AI Kết nối Endpoint Tùy biến & Streaming (AI Gateway & Streaming)](#trụ-cột-2-engine-ai-kết-nối-endpoint-tùy-biến--streaming-ai-gateway--streaming)
   - [Trụ cột 3: Bộ Công cụ Biên tập & Trợ lý Viết văn (Smart Editorial Suite)](#trụ-cột-3-bộ-công-cụ-biên-tập--trợ-lý-viết-văn-smart-editorial-suite)
   - [Trụ cột 4: Cổng Kết nối AI Agent Ngoài & MCP Server (Headless Bridge / MCP)](#trụ-cột-4-cổng-kết-nối-ai-agent-ngoài--mcp-server-headless-bridge--mcp)
   - [Trụ cột 5: Triển khai Sideload & Giao diện Chuẩn Hóa (Sideloading & UI)](#trụ-cột-5-triển-khai-sideload--giao-diện-chuẩn-hóa-sideloading--ui)
4. [BẢNG CHECKLIST TÍNH NĂNG ĐỐI CHIẾU (VERIFICATION MATRIX)](#4-bảng-checklist-tính-năng-đối-chiếu-verification-matrix)

---

# 1. TỔNG HỢP VÀ ĐỐI CHIẾU CÁC REPO CỐT LÕI

| Repo Tham Khảo | Điểm Sáng / Tính Năng Hay Nhất Bê Về | Đóng Góp Vào Hệ Thống Mới |
| :--- | :--- | :--- |
| **`claude-word-addin`** | • Chế độ **Track Changes (Redlines)** bằng OOXML chuẩn.<br>• Phân tích chỉ số đọc hiểu văn bản.<br>• Streaming token-by-token mượt mà qua SSE. | Bộ hiển thị Track Changes & Trình phân tích văn bản. |
| **`vivword`** | • **Chèn Margin Comments (Bình luận lề trang)** bằng WordApi 1.4.<br>• **Hệ thống Skills** (Prompt thẻ 1-chạm: Sửa giọng văn, Cắt từ thừa...).<br>• **Voz da casa (House Voice)**: Căn chỉnh giọng văn cá nhân.<br>• **Lưu nhớ theo file Word (`document.settings`)**: mỗi file nhớ lịch sử chat riêng. | UX Biên tập & Nhận xét thông minh, bộ nhớ theo từng tài liệu. |
| **`office-copilot-lite`** | • Tương thích Gateway linh hoạt.<br>• Cấu hình endpoint đơn giản, gọn nhẹ. | Thiết kế kết nối AI Gateway đa năng. |
| **`Claude-Word-Plugin`** | • **Context Attachment Chips**: Gắn tag đính kèm vùng chọn, trích dẫn câu trả lời cũ, link web.<br>• Hỗ trợ chế độ Thinking Effort (Suy luận sâu). | Cơ chế đính kèm ngữ cảnh trực quan & Thinking Mode. |
| **`wordbridge`** | • **Hệ thống Tools & API hoàn chỉnh cho AI Agent / MCP** (Scout paragraphs, exact anchor search, delete, set style, insert OOXML, review changes). | Cổng kết nối Agent ngoài & MCP Server điều khiển Word trực tiếp. |

---

# 2. KIẾN TRÚC TỔNG THỂ HỆ THỐNG

Hệ thống hoạt động với kiến trúc **Dual-Mode (2 Chế độ Song Song)** thông qua **AI Backend / Bridge Server**:

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

# 3. DANH MỤC TÍNH NĂNG CHI TIẾT (FEATURE SPECIFICATION)

---

## Trụ cột 1: Tương tác Word Chuyên sâu & Điều hướng (Word-Native & Navigation)

* [x] **1.1. Điều hướng Nhảy đoạn tương tác (Interactive Click-to-Jump / Anchor Navigation):**
  * Đã triển khai thuật toán 3 lớp định vị trong `public/modules/word-bridge.js` (`selectRange`) và render thẻ chip tự động trong `public/app.js`. Nhấp vào chip sẽ cuộn màn hình và bôi đen câu văn trong Word.
* [x] **1.2. Bình luận Lề Trang (Margin Comments Injection):**
  * Đã tích hợp `range.insertComment(commentText)` chuẩn WordApi 1.4 trong `public/modules/word-bridge.js` và nút bấm `💬 Bình luận lề`.
* [x] **1.3. Chế độ Theo dõi Sửa đổi (Track Changes / Redlines):**
  * Đã tích hợp hybrid OOXML `<w:ins>` / `<w:del>` trong `buildTrackedChangeOoxml` và `insertWithTrackChanges` hiển thị chuẩn vạch đỏ trên mọi nền tảng Word.
* [x] **1.4. Đa dạng Chế độ Chèn Nội dung (Flexible Insertion Modes):**
  * Đã tích hợp 4 nút hành động dưới mỗi câu trả lời: 🔄 *Thay thế*, ➕ *Tại con trỏ*, ⬇️ *Sau đoạn*, 📄 *Cuối trang*.
* [x] **1.5. Đọc & Bắt Ngữ cảnh Linh hoạt (Smart Context Capture):**
  * Đã tích hợp nút `🎯 Bôi đen` và `📑 Toàn bài` kèm đếm số từ và thanh chip ngữ cảnh có nút xóa.
* [x] **1.6. Phân tích Thống kê & Độ Dễ Đọc (Text Analytics):**
  * Đã tích hợp `getTextAnalytics` tính số từ, số ký tự và thời gian đọc ước tính.

---

## Trụ cột 2: Engine AI Kết nối Endpoint Tùy biến & Streaming (AI Gateway & Streaming)

* [x] **2.1. Phản hồi Thời gian Thực (Smooth Token-by-Token Streaming):**
  * Đã triển khai luồng SSE mượt mà trong `server/ai-proxy.js` và `public/modules/sse-client.js` kèm nút ⏹️ *Dừng sinh (AbortController)*.
* [x] **2.2. Hỗ trợ Kết nối AI Endpoint Tùy biến (Custom AI Endpoint Integration):**
  * Tương thích hoàn toàn cả chuẩn OpenAI `/v1/chat/completions` (DeepSeek, OpenAI, Gemini) và Anthropic `/v1/messages`. Cấu hình dễ dàng qua Modal Cài đặt hoặc `.env`.
* [x] **2.3. Hỗ trợ Chế độ Suy Luận Sâu (Thinking Mode / Extended Thinking):**
  * Đã bóc tách thẻ `<think>` / `reasoning_content` / `thinking_delta` và render thành Accordion gấp gọn `🧠 Quá trình suy luận`.

---

## Trụ cột 3: Bộ Công cụ Biên tập & Trợ lý Viết văn (Smart Editorial Suite)

* [x] **3.1. Hệ thống Kỹ năng Nhanh 1-Chạm (Skills System):**
  * Đã cấu hình 8 kỹ năng biên tập mặc định trong `public/skills.json` (Viết lại mượt, Sửa chính tả, Trang trọng hóa, Gọt từ thừa, Tóm tắt ý, Dịch song ngữ, Nhận xét lề trang, Sửa văn phong AI).
* [x] **3.2. Giọng Văn Tham Chiếu (House Voice / Căn Chỉnh Phong Cách):**
  * Đã tích hợp trường dán đoạn mẫu trong Cài đặt và tự động chèn vào thẻ `<personal_writing_style>` của System Prompt.
* [x] **3.3. Lưu Nhớ Trạng Thái Theo Từng File Word (`document.settings` Persistence):**
  * Đã triển khai `public/modules/doc-state.js` lưu nhớ lịch sử chat và cài đặt trực tiếp vào metadata của file `.docx` với debounced autosave.
* [x] **3.4. Đính kèm Ngữ cảnh Đa dạng (Context Attachment Chips):**
  * Đã hiển thị thanh Context Chip trên đầu khung chat với số từ đếm và nút xóa.

---

## Trụ cột 4: Cổng Kết nối AI Agent Ngoài & MCP Server (Headless Bridge / MCP)

* [x] **4.1. Cổng Kết nối Đa Giao Thức (Multi-Protocol Bridge):**
  * WebSocket Bridge (`/ws`), REST API (`/op`, `/ops`, `/tools`, `/api/status`) và MCP Server qua `stdio`.
* [x] **4.2. Bộ Danh mục Tools Đầy đủ cho Agent (Agent Tool Catalog):**
  * Đã triển khai 14+ Tools trong `server/tools.js`, `server/mcp-server.js` và `public/modules/ws-client.js`.
* [x] **4.3. Cơ chế Khóa & Báo cáo Lỗi An toàn (Safe Execution & Error Serialization):**
  * Hiển thị banner trạng thái `🤖 AI Agent đang thao tác...` khi Agent ngoài chạy lệnh.

---

## Trụ cột 5: Triển khai Sideload & Giao diện Chuẩn Hóa (Sideloading & UI)

* [x] **5.1. Phương thức Cài đặt Chuẩn (Standard Public Sideloading):**
  * File `manifest.xml` đã được xác thực 100% hợp lệ (`npx office-addin-manifest validate manifest.xml`).
* [x] **5.2. Giao diện Đẹp, Hiện đại & Hoàn toàn bằng Tiếng Việt:**
  * Thiết kế hiện đại trong `public/styles.css` và `public/index.html` với Dark/Light mode, font chữ Inter sắc nét.

---

# 4. BẢNG CHECKLIST TÍNH NĂNG ĐỐI CHIẾU (VERIFICATION MATRIX)

| Mã | Nhóm Tính Năng | Tên Tính Năng Cụ Thể | Trạng Thái | Mức Độ Ưu Tiên | Nguồn Tinh Hoa |
| :---: | :--- | :--- | :---: | :---: | :---: |
| **W-01** | Tương tác Word | **Click-to-Jump**: AI chỉ đoạn nào, click nhảy & bôi đen đoạn đó trong Word | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `word-bridge.js` + `app.js` |
| **W-02** | Tương tác Word | **Margin Comments**: Chèn nhận xét lề phải Word không làm hỏng bài gốc | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `vivword` (`insertComment`) |
| **W-03** | Tương tác Word | **Track Changes**: Ghi nhận chi tiết phần AI sửa đổi (Redlines) | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `claude-word-addin` (OOXML) |
| **W-04** | Tương tác Word | **4 chế độ chèn**: Thay thế vùng chọn, Chèn con trỏ, Chèn sau, Chèn cuối | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `word-bridge.js` + Actions Bar |
| **W-05** | Tương tác Word | **Smart Context**: Đọc nhanh vùng chọn hoặc toàn bộ tài liệu | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `word-bridge.js` |
| **AI-01** | AI Engine | **Streaming Token-by-Token**: Chữ chạy mượt mà theo thời gian thực (SSE) | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `ai-proxy.js` + `sse-client.js` |
| **AI-02** | AI Engine | **Custom Endpoint**: Tích hợp AI Endpoint có sẵn của bạn | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `ai-proxy.js` + Settings Modal |
| **AI-03** | AI Engine | **Thinking Mode**: Hỗ trợ hiển thị khối suy luận sâu của AI (R1 / Sonnet) | ✅ Hoàn thành | ⚡ P1 (Nâng cao) | `thinking-box` accordion |
| **ED-01** | Biên tập | **Hệ thống Skills 1-Chạm**: Viết lại mượt mà, Sửa chính tả, Tóm tắt, Dịch... | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `skills.json` + `skillsBar` |
| **ED-02** | Biên tập | **House Voice**: Cung cấp đoạn văn mẫu để AI học chuẩn giọng văn cá nhân | ✅ Hoàn thành | ⚡ P1 (Nâng cao) | `doc-state.js` + `<personal_writing_style>` |
| **ED-03** | Biên tập | **Lưu nhớ theo file Word**: Tự nhớ lịch sử chat theo từng file `.docx` | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `doc-state.js` (`document.settings`) |
| **ED-04** | Biên tập | **Context Chips**: Gắn thẻ đính kèm trích dẫn văn bản trực quan | ✅ Hoàn thành | ⚡ P1 (Nâng cao) | `context-bar` |
| **AG-01** | Agent Bridge | **Cổng WebSocket & REST API**: Kết nối 2 chiều đồng bộ với Word | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `server.js` (`/ws`, `/op`, `/ops`) |
| **AG-02** | Agent Bridge | **Chuẩn MCP Server**: Cho phép Claude Code, Cursor... tự động nhận diện Word | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `mcp-server.js` (StdioServerTransport) |
| **AG-03** | Agent Bridge | **Bộ 14+ Tools đầy đủ**: Đọc, sửa, nhảy con trỏ, chèn OOXML, style, comment | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `tools.js` + `ws-client.js` |
| **DP-01** | Cài đặt | **Sideloading chuẩn hóa**: File `manifest.xml` cài đặt đơn giản trên Web/Desktop | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `manifest.xml` (Validated) |
| **UI-01** | Giao diện | **UI Tiếng Việt hiện đại**: Tối giản, thẩm mỹ cao, hỗ trợ Dark/Light mode | ✅ Hoàn thành | 🔥 P0 (Cốt lõi) | `index.html` + `styles.css` |
