# TÀI LIỆU THIẾT KẾ KIẾN TRÚC & GIẢI PHÁP KỸ THUẬT CHUYÊN SÂU
## DỰ ÁN: AI WORD ADD-IN & MCP AGENT BRIDGE

> **Mục tiêu tài liệu:** Tổng hợp kết quả Brainstorming đa chiều qua nhiều vòng (Multi-round Brainstorming) để giải quyết triệt để tất cả các bài toán kỹ thuật, bẫy tiềm ẩn (gotchas) của Office.js, thiết kế giao diện UI/UX trực quan nhất cho người dùng, và xây dựng chuẩn giao tiếp MCP hoàn chỉnh cho AI Agent bên ngoài.

---

# MỤC LỤC
1. [VÒNG 1: THIẾT KẾ UX/UI & CƠ CHẾ TƯƠNG TÁC WORD CHUYÊN SÂU](#vòng-1-thiết-kế-uxui--cơ-chế-tương-tác-word-chuyên-sâu)
   - [1.1. Cơ chế Điều hướng Nhấp-để-Nhảy (Click-to-Jump / Anchor Navigation)](#11-cơ-chế-điều-hướng-nhấp-để-nhảy-click-to-jump--anchor-navigation)
   - [1.2. Cơ chế Bình luận Lề Trang (Margin Comments Injection)](#12-cơ-chế-bình-luận-lề-trang-margin-comments-injection)
   - [1.3. Cơ chế Track Changes / Redlines Xuyên Nền Tảng](#13-cơ-chế-track-changes--redlines-xuyên-nền-tảng)
   - [1.4. Quản lý Vùng Chọn (Selection Preservation) & 4 Chế độ Chèn](#14-quản-lý-vùng-chọn-selection-preservation--4-chế-độ-chèn)
   - [1.5. Hệ thống Skills 1-Chạm & Đồng bộ Giọng văn (House Voice)](#15-hệ-thống-skills-1-chạm--đồng-bộ-giọng-văn-house-voice)
   - [1.6. Cơ chế Lưu trạng thái Gắn liền File Word (`document.settings`)](#16-cơ-chế-lưu-trạng-thái-gắn-liền-file-word-documentsettings)
2. [VÒNG 2: THIẾT KẾ AI ENGINE, GATEWAY ENDPOINT & STREAMING](#vòng-2-thiết-kế-ai-engine-gateway-endpoint--streaming)
   - [2.1. Kiến trúc Kết nối AI Endpoint Linh hoạt](#21-kiến-trúc-kết-nối-ai-endpoint-linh-hoạt)
   - [2.2. Pipeline Truyền dẫn Streaming SSE Token-by-Token](#22-pipeline-truyền-dẫn-streaming-sse-token-by-token)
   - [2.3. Xử lý Chế độ Suy nghĩ Sâu (Thinking Mode / DeepSeek R1 / Claude 3.7)](#23-xử-lý-chế-độ-suy-nghĩ-sâu-thinking-mode--deepseek-r1--claude-37)
   - [2.4. Quy trình Lắp ráp Ngữ cảnh (Context Assembly Pipeline)](#24-quy-trình-lắp-ráp-ngữ-cảnh-context-assembly-pipeline)
3. [VÒNG 3: THIẾT KẾ CỔNG MCP SERVER & DUAL-MODE BRIDGE](#vòng-3-thiết-kế-cổng-mcp-server--dual-mode-bridge)
   - [3.1. Mô hình Đồng bộ 2 Chiều (Dual-Mode Sync Architecture)](#31-mô-hình-đồng-bộ-2-chiều-dual-mode-sync-architecture)
   - [3.2. Đặc tả Giao thức MCP (Model Context Protocol) & Danh mục 14+ Tools](#32-đặc-tả-giao-thức-mcp-model-context-protocol--danh-mục-14-tools)
   - [3.3. Xử lý Tranh chấp (Concurrency & State Locking)](#33-xử-lý-tranh-chấp-concurrency--state-locking)
4. [VÒNG 4: GIẢI PHÁP ỔN ĐỊNH, OFFICE.JS QUIRKS & XỬ LÝ NGOẠI LỆ](#vòng-4-giải-pháp-ổn-định-officejs-quirks--xử-lý-ngoại-lệ)
   - [4.1. Khắc phục Lỗi Thường Gặp trên Office.js](#41-khắc-phục-lỗi-thường-gặp-trên-officejs)
   - [4.2. Xử lý Tài liệu Lớn (> 50.000 từ)](#42-xử-lý-tài-liệu-lớn--50000-từ)
   - [4.3. Kiến trúc Cấu trúc Thư mục Dự án Tinh gọn](#43-kiến-trúc-cấu-trúc-thư-mục-dự-án-tinh-gọn)

---

# VÒNG 1: THIẾT KẾ UX/UI & CƠ CHẾ TƯƠNG TÁC WORD CHUYÊN SÂU

---

## 1.1. Cơ chế Điều hướng Nhấp-để-Nhảy (Click-to-Jump / Anchor Navigation)

### Bài toán:
Khi AI trả lời: *"Đoạn 2 của bạn có câu 'Chúng tôi xin cam kết...' dùng từ chưa chuẩn, nên sửa thành..."*, làm sao để người dùng bấm vào câu trích dẫn trên Sidebar thì Word tự động **cuộn tới đúng trang đó và bôi đen (highlight) chính xác câu văn đó**?

```
┌───────────────────────────────────────┐         ┌───────────────────────────────────────┐
│          SIDEBAR TRONG WORD           │         │         TÀI LIỆU WORD CHÍNH           │
├───────────────────────────────────────┤         ├───────────────────────────────────────┤
│ [AI]: Đoạn sau cần diễn đạt lại:      │         │ Trang 3:                              │
│ ┌───────────────────────────────────┐ │  Click  │                                       │
│ │ 🎯 "Chúng tôi xin cam kết..."     │─┼────────►│ ...                                   │
│ └───────────────────────────────────┘ │  Nhảy   │ [██ Chúng tôi xin cam kết... ██]     │
│ Gợi ý: "Chúng tôi cam đoan..."       │         │ (Tự động cuộn & bôi đen câu này)      │
└───────────────────────────────────────┘         └───────────────────────────────────────┘
```

### Giải pháp kỹ thuật chuyên sâu:
1. **Trích xuất Anchor tự động:**
   * Trong phản hồi của AI, hệ thống tự động nhận diện các đoạn trích dẫn (nằm trong dấu ngoặc kép `""` hoặc cú pháp `> trích dẫn` hoặc thẻ metadata `[[anchor: đoạn văn bản]]`).
   * Render thành các thẻ chip tương tác có biểu tượng `🎯 [Xem trong văn bản]`.
2. **Thuật toán Định vị 3 Lớp (3-Tier Locator Strategy):**
   * **Lớp 1 (Exact Match):** Sử dụng `context.document.body.search(anchorText, { matchCase: false })`. Nếu tìm thấy chính xác 1 vị trí -> gọi `range.select()` ngay lập tức.
   * **Lớp 2 (Fuzzy / Normalized Match):** Nếu Lớp 1 không khớp (do khác biệt khoảng trắng `\r\n`, tab hay dấu ngoặc), hệ thống chuẩn hóa chuỗi (bỏ khoảng trắng thừa) và so khớp đoạn đầu (prefix 30 ký tự) + đoạn cuối (suffix 30 ký tự).
   * **Lớp 3 (Paragraph Index Fallback):** Nếu neo bằng vị trí đoạn (`paragraphIndex`), hệ thống lấy `context.document.body.paragraphs.getItemAt(index).select()`.
3. **Hiệu ứng thị giác (Visual Feedback):**
   * Sau khi gọi `targetRange.select()`, Word lập tức scroll khung nhìn đến vùng chọn và bôi đen màu xám/xanh bản địa của Word.
   * Trên UI Taskpane, thẻ chip sáng hiệu ứng `pulse` xanh lá báo hiệu đã định vị thành công.

---

## 1.2. Cơ chế Bình luận Lề Trang (Margin Comments Injection)

### Bài toán:
Nhiều trường hợp người dùng (ví dụ: chị bạn hoặc giáo viên, biên tập viên) chỉ muốn **AI để lại lời nhận xét/góp ý bên lề trang** như một người chấm bài, chứ không muốn AI ghi đè làm xáo trộn văn bản đang viết dở.

### Giải pháp kỹ thuật:
1. **Sử dụng API `WordApi 1.4`:**
   ```javascript
   await Word.run(async (context) => {
     const selection = context.document.getSelection();
     selection.insertComment(commentText);
     await context.sync();
   });
   ```
2. **Fallback an toàn cho các môi trường Word cũ:**
   * Hệ thống kiểm tra tính năng `Office.context.requirements.isSetSupported('WordApi', '1.4')`.
   * Nếu môi trường hỗ trợ: Hiện nút **💬 Thêm nhận xét lề trang**.
   * Nếu môi trường cũ (không hỗ trợ comment qua API): Tự động chuyển đổi sang chế độ chèn đoạn ghi chú trong ngoặc vuông `[AI Ghi chú: ...]` ngay dưới đoạn văn.

---

## 1.3. Cơ chế Track Changes / Redlines Xuyên Nền Tảng

### Bài toán:
Trên một số phiên bản Word (đặc biệt là Word for Mac), lệnh bật Track Changes bằng mã JS `document.changeTrackingMode = "TrackAll"` thường bị lỗi API của Microsoft (vấn đề đã được ghi nhận tại GitHub OfficeDev #2797, #6246).

### Giải pháp kỹ thuật (Hybrid Track Changes Engine):
1. **Tầng 1 (Native Tracking):** Cố gắng gọi `context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll`.
2. **Tầng 2 (Literal OOXML Injection - Chắc chắn 100%):**
   * Khi AI sửa văn bản, thay vì chèn chuỗi thô, hệ thống sinh đoạn mã OpenXML (OOXML) chứa cấu trúc `<w:ins>` (thêm mới) và `<w:del>` (xóa bỏ) có gắn tên tác giả `AI Assistant`:
   ```xml
   <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
     <w:del w:id="1" w:author="AI Assistant" w:date="2026-08-31T12:00:00Z">
       <w:r><w:delText>từ cũ bị xóa</w:delText></w:r>
     </w:del>
     <w:ins w:id="2" w:author="AI Assistant" w:date="2026-08-31T12:00:00Z">
       <w:r><w:t>từ mới được thêm vào</w:t></w:r>
     </w:ins>
   </w:p>
   ```
   * Chèn bằng `selection.insertOoxml(ooxmlString, Word.InsertLocation.replace)`.
   * **Kết quả:** Hiển thị chuẩn xác vạch đỏ Redlines (gạch ngang chữ cũ, gạch chân chữ mới) trên mọi phiên bản Word mà không cần quan tâm Word có đang bật chế độ Track Changes hay không!

---

## 1.4. Quản lý Vùng Chọn (Selection Preservation) & 4 Chế độ Chèn

### Bài toán:
Khi người dùng bôi đen 1 đoạn văn trong Word rồi bấm chuột sang Taskpane (sidebar), Word có thể bị mất tiêu điểm (blur) hoặc con trỏ bị nhảy sang vị trí khác nếu người dùng lỡ bấm nhầm.

### Giải pháp:
1. **Lưu cache vùng chọn (Selection Snapshot):**
   * Mỗi khi người dùng bấm nút *"Đọc vùng chọn"* hoặc chọn 1 Kỹ năng (Skill), hệ thống lập tức lưu lại tọa độ/văn bản của vùng chọn vào bộ nhớ tạm `state.cachedSelectionRange`.
2. **4 Chế độ Chèn Nội dung Trực quan:**
   * 🔄 **Thay thế (Replace):** Thay thế chính xác vùng văn bản đã chọn ban đầu.
   * ➕ **Tại con trỏ (At Cursor):** Chèn vào vị trí con trỏ hiện tại.
   * ⬇️ **Sau đoạn chọn (Insert After):** Tạo một đoạn văn mới ngay sau đoạn vừa sửa.
   * 📄 **Cuối tài liệu (Append End):** Chèn xuống cuối trang tài liệu.
3. **Thanh Actions Bar dưới mỗi câu trả lời của AI:**
   * Mỗi câu trả lời của AI trên chat sẽ có sẵn 4 icon nút bấm thao tác nhanh để người dùng chọn cách chèn phù hợp nhất chỉ với 1 click.

---

## 1.5. Hệ thống Skills 1-Chạm & Đồng bộ Giọng văn (House Voice)

### A. Hệ thống Kỹ năng 1-Chạm (Skills System)
* Được thiết kế dạng thanh thẻ Pill trượt ngang ở đầu Taskpane.
* Danh mục kỹ năng mặc định được định nghĩa dưới dạng JSON có cấu trúc rõ ràng:
  ```json
  [
    {
      "id": "polish",
      "label": "✍️ Viết lại mượt",
      "prompt": "Hãy viết lại đoạn văn sau cho trôi chảy, tự nhiên, văn phong mạch lạc và lôi cuốn hơn nhưng giữ nguyên nội dung cốt lõi.",
      "mode": "replace"
    },
    {
      "id": "grammar",
      "label": "🔍 Sửa ngữ pháp & chính tả",
      "prompt": "Hãy tìm và sửa tất cả các lỗi chính tả, dấu câu và ngữ pháp tiếng Việt trong đoạn văn sau. Chỉ rõ vị trí đã sửa.",
      "mode": "replace"
    },
    {
      "id": "formal",
      "label": "👔 Trang trọng hóa",
      "prompt": "Hãy chuyển đổi đoạn văn sau sang văn phong hành chính, trang trọng, chuẩn mực công vụ.",
      "mode": "replace"
    },
    {
      "id": "cut_fluff",
      "label": "✂️ Gọt từ thừa",
      "prompt": "Hãy cô đọng đoạn văn sau, loại bỏ các từ ngữ dư thừa, sáo rỗng, giảm 20-30% độ dài nhưng giữ nguyên toàn bộ ý chính.",
      "mode": "replace"
    },
    {
      "id": "summary",
      "label": "📝 Tóm tắt ý",
      "prompt": "Hãy tóm tắt các luận điểm cốt lõi của văn bản dưới dạng các gạch đầu dòng rõ ràng, súc tích.",
      "mode": "insert_after"
    },
    {
      "id": "comment_review",
      "label": "💬 Nhận xét biên tập",
      "prompt": "Đóng vai trò là một biên tập viên chuyên nghiệp, hãy đưa ra nhận xét, đánh giá ưu/nhược điểm và gợi ý cải thiện cho đoạn văn sau.",
      "mode": "comment"
    }
  ]
  ```

### B. Giọng văn tham chiếu (House Voice / Personal Anchor)
* Người dùng có thể mở tab Cài đặt và dán 1-3 đoạn văn mẫu chuẩn phong cách mình thích.
* **Cơ chế hoạt động:** Hệ thống tự động gói đoạn văn này vào system prompt dưới thẻ:
  ```
  <personal_writing_style>
  [Đoạn văn mẫu của người dùng]
  Quy tắc: Luôn mô phỏng đúng nhịp điệu câu, cách dùng từ và phong thái của đoạn văn mẫu trên trong mọi câu trả lời.
  </personal_writing_style>
  ```

---

## 1.6. Cơ chế Lưu trạng thái Gắn liền File Word (`document.settings`)

### Bài toán:
Chị bạn đang làm dở tài liệu Hợp đồng A, đã chat với AI được 10 câu. Ngày mai mở lại Hợp đồng A thì lịch sử chat có còn không? Nếu mở tài liệu B thì có bị lẫn lịch sử của tài liệu A không?

### Giải pháp:
* Sử dụng API `Office.context.document.settings`: Dữ liệu được lưu **trực tiếp vào bên trong metadata của chính file `.docx`**.
* **Dữ liệu được lưu trữ gồm:**
  1. `chat_history`: Danh sách tin nhắn hỏi đáp của tài liệu.
  2. `document_voice`: Mẫu giọng văn thiết lập riêng cho tài liệu đó.
  3. `custom_system_prompt`: Hướng dẫn riêng cho tài liệu (nếu có).
* **Cơ chế Auto-save Debounce:** Sau mỗi lượt chat, hệ thống tự động lưu ngầm (`saveAsync`) sau 500ms mà không làm gián đoạn trải nghiệm gõ phím.

---

# VÒNG 2: THIẾT KẾ AI ENGINE, GATEWAY ENDPOINT & STREAMING

---

## 2.1. Kiến trúc Kết nối AI Endpoint Linh hoạt

Backend Server đóng vai trò là một **Proxy Adapter** siêu nhẹ, tiếp nhận yêu cầu từ Taskpane và chuyển tiếp đến AI API Endpoint của bạn:

```
┌───────────────────────────┐                  ┌───────────────────────────┐                  ┌───────────────────────────┐
│     TASKPANE WEBVIEW      │  POST /api/chat  │       BACKEND SERVER      │   Forward Req    │   YOUR AI API ENDPOINT    │
│  (Word Sidebar UI)        ├─────────────────►│   (Node.js / Express)     ├─────────────────►│   (OpenAI / Anthropic /   │
│                           │◄─────────────────┤                           │◄─────────────────┤    Custom LLM Gateway)    │
│                           │    SSE Stream    │                           │    SSE Stream    │                           │
└───────────────────────────┘                  └───────────────────────────┘                  └───────────────────────────┘
```

* **Ưu điểm:**
  1. Không bao giờ lo lỗi CORS (Cross-Origin Resource Sharing) từ trình duyệt webview của Office.
  2. Không cần phơi bày API Key ở phía giao diện người dùng.
  3. Tương thích linh hoạt cả định dạng payload OpenAI (`/v1/chat/completions`) và Anthropic (`/v1/messages`).

---

## 2.2. Pipeline Truyền dẫn Streaming SSE Token-by-Token

1. Phía Server mở kết nối `Content-Type: text/event-stream`.
2. Đọc luồng dữ liệu `upstream.body.getReader()` từ AI Endpoint và đẩy thẳng từng chunk `data: {"delta": "..."}` về Taskpane.
3. Phía Taskpane sử dụng trình phân giải SSE (Server-Sent Events Parser) để nối chữ trực tiếp vào DOM theo thời gian thực:
   * Chữ xuất hiện tức thì với tốc độ sinh của mô hình.
   * Tự động cuộn xuống cuối khung chat (`chatContainer.scrollTop = chatContainer.scrollHeight`).
   * Hỗ trợ nút ⏹️ **Dừng sinh (AbortController)** để ngắt kết nối ngay lập tức khi người dùng muốn dừng lại.

---

## 2.3. Xử lý Chế độ Suy nghĩ Sâu (Thinking Mode / DeepSeek R1 / Claude 3.7)

Khi sử dụng các mô hình suy luận sâu (Reasoning Models) như DeepSeek R1 hay Claude 3.7 Sonnet:
* AI sẽ trả về khối suy nghĩ trước khi đưa ra câu trả lời chính thức (`<think>...</think>` hoặc `reasoning_content`).
* **Thiết kế UI:**
  * Khối suy nghĩ được bao bọc trong một Accordion gấp gọn màu xám nhạt với icon `🧠 Quá trình suy luận (Nhấp để mở)`.
  * Trong lúc AI đang suy nghĩ, icon sẽ xoay nhẹ (thinking animation).
  * Khi bắt đầu sang câu trả lời chính, khối suy nghĩ tự động thu nhỏ lại để không chiếm diện tích màn hình của người dùng.

---

## 2.4. Quy trình Lắp ráp Ngữ cảnh (Context Assembly Pipeline)

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                               CONTEXT ASSEMBLY PIPELINE                                │
├──────────────────────────────┬─────────────────────────────────────────────────────────┤
│ 1. System Prompt Base        │ "Bạn là trợ lý AI chuyên nghiệp hỗ trợ soạn thảo Word..."│
│ 2. House Voice (Nếu có)      │ "<personal_writing_style>...</personal_writing_style>"   │
│ 3. Active Context            │ [Đoạn văn bôi đen] HOẶC [Tóm lược toàn bộ tài liệu]     │
│ 4. Selected Skill Template   │ "Hãy viết lại đoạn văn sau cho trang trọng hơn: ..."    │
│ 5. Conversation History      │ 4-6 lượt hội thoại gần nhất trong file Word hiện tại    │
│ 6. User Prompt               │ Yêu cầu bổ sung của người dùng                          │
└──────────────────────────────┴─────────────────────────────────────────────────────────┘
```

---

# VÒNG 3: THIẾT KẾ CỔNG MCP SERVER & DUAL-MODE BRIDGE

---

## 3.1. Mô hình Đồng bộ 2 Chiều (Dual-Mode Sync Architecture)

Hệ thống cung cấp **Chế độ Kép (Dual-Mode)** chạy song song trên cùng 1 server:
1. **Chế độ Giao diện Taskpane (Cho người dùng):** Hoạt động qua HTTP/SSE phục vụ người dùng tương tác trực tiếp bằng chuột.
2. **Chế độ Agent Ngoài & MCP Server (Cho AI Agent):** Hoạt động qua cổng WebSocket (`ws://localhost:3001/ws`) kết nối trực tiếp với Taskpane Office.js. Khi Claude Code, Cursor hay Codex gửi lệnh qua MCP Server, server sẽ đẩy lệnh qua WebSocket để Word thực thi tức thì.

```
                    ┌──────────────────────────────────────────────┐
                    │            EXTERNAL AI AGENTS                │
                    │   (Claude Code, Cursor, Codex, OpenCode)     │
                    └──────────────────────┬───────────────────────┘
                                           │ MCP Protocol (stdio / JSON-RPC)
                                           ▼
                    ┌──────────────────────────────────────────────┐
                    │               LOCAL MCP SERVER               │
                    │            (Node.js / Express Bridge)        │
                    └──────────────────────┬───────────────────────┘
                                           │ WebSocket Push (/ws)
                                           ▼
                    ┌──────────────────────────────────────────────┐
                    │         WORD TASKPANE OFFICE.JS RUNTIME      │
                    │         (Tự động thực thi lệnh trên Word)    │
                    └──────────────────────────────────────────────┘
```

---

## 3.2. Đặc tả Giao thức MCP & Danh mục 14+ Tools

Server triển khai chuẩn **Model Context Protocol (MCP)** của Anthropic. Khi kết nối, Agent ngoài sẽ tự động nhận diện danh mục tools phong phú:

| Tool Name | Tham số đầu vào (JSON Schema) | Chức năng chi tiết |
| :--- | :--- | :--- |
| `word_ping` | `{}` | Kiểm tra trạng thái kết nối tới Word Taskpane. |
| `word_getText` | `{ limit?: number }` | Đọc toàn bộ văn bản thô trong tài liệu Word. |
| `word_getParagraphs` | `{ styleFilter?: string, limit?: number, textLimit?: number }` | Đọc danh sách đoạn văn có cấu trúc kèm kiểu định dạng (Heading, Normal...). |
| `word_getSelection` | `{}` | Đọc đoạn văn bản người dùng đang bôi đen trên màn hình Word. |
| `word_selectRange` | `{ anchor?: string, paragraphIndex?: number }` | **Điều khiển Word tự cuộn và bôi đen đoạn văn bản chỉ định.** |
| `word_findReplace` | `{ find: string, replace: string, matchCase?: boolean, maxReplacements?: number }` | Tìm kiếm và thay thế chuỗi trên toàn bộ tài liệu. |
| `word_insertAfterText`| `{ anchor: string, text: string, asParagraph?: boolean, style?: string }` | Tìm từ khóa `anchor` và chèn nội dung mới ngay phía sau. |
| `word_deleteText` | `{ find: string, maxDeletions?: number }` | Tìm và xóa đoạn văn bản chỉ định. |
| `word_setParagraphStyle`| `{ anchor: string, style: string }` | Đổi định dạng kiểu đoạn văn (Heading 1, Heading 2, Quote, Normal). |
| `word_insertComment`| `{ anchor: string, commentText: string }` | **Tạo bình luận lề phải Word tại đoạn văn bản chỉ định.** |
| `word_insertOoxml` | `{ anchor: string, ooxml: string, location: "before"\|"after"\|"replace" }` | Chèn mã OOXML tùy biến (bảng biểu, cấu trúc nâng cao, `<w:ins>`). |
| `word_getTrackedChanges`| `{ timeoutMs?: number }` | Liệt kê tất cả các sửa đổi đang chờ duyệt (Insertions, Deletions, Formats). |
| `word_setTrackChanges`| `{ on: boolean }` | Bật/tắt chế độ ghi nhận sửa đổi trong Word. |
| `word_getMetadata` | `{}` | Đọc thông tin thuộc tính file (Title, Author, Subject, Last Author). |

---

## 3.3. Xử lý Tranh chấp (Concurrency & State Locking)

* Khi Agent ngoài đang thực hiện một chuỗi thao tác (Batch Ops) trên tài liệu:
  * Taskpane sẽ hiển thị một biểu tượng ổ khóa nhẹ `🔒 Agent đang thao tác...` để người dùng không bấm nhầm nút trên giao diện.
  * Khi Agent hoàn tất lệnh, giao diện mở khóa tức thì và phát thông báo `✅ Đã đồng bộ thay đổi từ Agent`.

---

# VÒNG 4: GIẢI PHÁP ỔN ĐỊNH, OFFICE.JS QUIRKS & XỬ LÝ NGOẠI LỆ

---

## 4.1. Khắc phục Lỗi Thường Gặp trên Office.js (Bẫy kinh nghiệm từ .ref)

1. **Lỗi `PropertyNotLoaded`:**
   * Trong Office.js, thuộc tính như `range.text` hay `body.paragraphs` không có sẵn cho đến khi gọi `.load()` và `context.sync()`.
   * **Quy tắc chuẩn:** Luôn tải đúng các trường cần thiết: `range.load("text"); await context.sync();`.
2. **Lỗi ngắt kết nối WebSocket khi ẩn Taskpane:**
   * Khi người dùng thu nhỏ Taskpane, trình duyệt Edge WebView2 có thể tạm dừng luồng WebSocket.
   * **Giải pháp:** Tự động kích hoạt cơ chế `Heartbeat (Ping/Pong)` mỗi 15 giây và tự động kết nối lại (`auto-reconnect`) sau 2 giây nếu bị mất kết nối.
3. **Lỗi mã hóa ký tự XML trong OOXML:**
   * Khi sinh chuỗi OOXML, các ký tự đặc biệt (`&`, `<`, `>`, `"`, `'`) phải được escape thành (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`) để tránh làm hỏng cấu trúc file Word.

---

## 4.2. Xử lý Tài liệu Lớn (> 50.000 từ)

* Nếu tài liệu quá dài (hàng trăm trang), việc gửi toàn bộ văn bản vào ngữ cảnh AI có thể gây tràn token và tốn thời gian.
* **Giải pháp:**
  * Đếm số từ trước khi nạp (`wordCount = text.split(/\s+/).length`).
  * Nếu `wordCount > 15.000 từ`: Hệ thống tự động chia thành danh mục tiêu đề (Outline/Headings) để người dùng chọn phần cần làm việc, hoặc chỉ nạp vùng chọn hiện tại.

---

## 4.3. Kiến trúc Cấu trúc Thư mục Dự án Tinh gọn

Cấu trúc mã nguồn chuẩn bị triển khai:

```
ai-add-in-word/
├── manifest.xml                # Office Add-in manifest chuẩn
├── package.json                # Dependencies tối giản (express, ws, @modelcontextprotocol/sdk)
├── server/                     # BACKEND, PROXY & MCP SERVER
│   ├── index.js                # Server chính khởi động HTTP & WebSocket
│   ├── ai-proxy.js             # SSE Proxy kết nối tới Custom AI Endpoint
│   ├── mcp-server.js           # Chuẩn Model Context Protocol Server
│   └── tools-catalog.js        # Khai báo schema 14+ Tools
├── public/                     # GIAO DIỆN TASKPANE
│   ├── index.html              # HTML giao diện Sidebar tiếng Việt
│   ├── styles.css              # Thiết kế giao diện Dark/Light mode cao cấp
│   ├── app.js                  # Bộ điều khiển giao diện chính
│   ├── skills.json             # Cấu hình danh mục kỹ năng 1-chạm
│   └── modules/                # MODULES XỬ LÝ CHUYÊN BIỆT
│       ├── word-bridge.js      # Click-to-Jump, Comments, OOXML, Selection
│       ├── doc-state.js        # Lưu nhớ chat & settings theo file Word
│       ├── sse-stream.js       # Phân giải luồng stream SSE
│       └── ws-client.js        # Lắng nghe lệnh điều khiển từ MCP Server
└── docs/
    ├── AI_WORD_ADDIN_FEATURE_SPEC.md     # Bản đặc tả tính năng & Checklist
    └── ARCHITECTURE_AND_SYSTEM_DESIGN.md # Tài liệu thiết kế kỹ thuật này
```
