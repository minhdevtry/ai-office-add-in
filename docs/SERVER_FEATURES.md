# Ori AI Server — Feature Inventory

> Thống kê toàn bộ chức năng server hiện có. Dùng làm reference khi thiết kế/refactor client.
> Cập nhật lần cuối: tự động từ source code `server/`.

---

## 1. Tổng quan

| Metric | Value |
|---|---|
| HTTP endpoints | **13** (1 SSE chat, 1 status, 1 tools catalog, 1 op, 1 ops, 6 license, 1 favicon, 1 manifest) |
| WebSocket endpoints | **1** (`/ws`) |
| MCP tools | **18** (stdio transport) |
| Default port | `3011` |
| Default host | `0.0.0.0` |
| HTTPS | optional (cần `office-addin-dev-certs`) |
| Auth | License (X-License-Key + X-Client-Id) + Supabase RPC |
| Rate limit | Chỉ áp dụng cho `POST /api/license/activate` (5 req/min/IP) |

---

## 2. HTTP Endpoints

### 2.1 Streaming & Core

#### `POST /api/chat` — AI Chat Proxy (SSE)
- **Auth**: required
- **Handler**: `server.js:90-100` → `ai-proxy.js:139-333`
- **Request body**:
  ```json
  {
    "messages": [{"role": "user|assistant|system", "content": "string|any"}],
    "systemPrompt": "string (optional)",
    "stream": true,
    "effort": "off|low|medium|high|xhigh (default: server config)",
    "maxTokens": "int (optional)",
    "temperature": "number (optional)"
  }
  ```
- **Response**: `text/event-stream` với 4 event types: `delta` (text), `thinking` (DeepSeek R1/Anthropic), `done` (kết thúc), `error` (lỗi giữa chừng)
- **Provider auto-detect**: `isAnthropicEndpoint()` check URL substring `anthropic` / `/messages` → dùng headers Anthropic; ngược lại → OpenAI-compatible (Bearer)

#### `GET /api/status` — System Health
- **Auth**: public
- **Response**:
  ```json
  {
    "status": "ok", "version": "1.0.0",
    "serverTime": "ISO",
    "aiEndpoint": "...", "aiModel": "...",
    "bridge": {"connectedCount": N, "clients": [...]}
  }
  ```

#### `GET /tools` — MCP Tool Catalog
- **Auth**: public
- Trả về `{name, version, tools: [...]}` — 18 tool definitions với inputSchema JSON

### 2.2 Word Tool Dispatch (qua WebSocket)

#### `POST /op` — Execute 1 op
- **Auth**: required
- **Body**: `{kind: "findReplace", find: "...", replace: "...", clientId?: "...", timeoutMs?: 15000}`
- **Response**: `{ok: true, result: ...}`

#### `POST /ops` — Execute batch ops
- **Auth**: required
- **Body**: mảng `[{kind, ...}, ...]` HOẶC `{ops: [...], stopOnError: bool}`
- **Response**: `{ok, total, completed, results: [...]}`

### 2.3 License API (`/api/license/*`)

| Method | Path | Auth | Purpose |
|---|---|---|---|
| `GET` | `/public-config` | public | `{aiModel, aiDefaultEffort, licenseRequired}` — không leak key |
| `POST` | `/activate` | public + rate-limit (5/min) | Enroll device, idempotent |
| `GET` | `/info` | required | Thông tin license + device hiện tại |
| `POST` | `/heartbeat` | required | Touch `last_seen_at`, trả `{ttl: 3600}` |
| `GET` | `/devices` | required | List active + revoked devices, mark `isCurrent` |
| `DELETE` | `/device/:deviceId` | required | Revoke 1 device (từ chối revoke cuối cùng) |

**Auth header**: `X-License-Key` + `X-Client-Id` (hoặc `X-Device-Id`).

**Activate body**:
```json
{ "email": "...", "licenseKey": "uuid", "deviceName": "..." (optional) }
```
**Response**:
```json
{ "ok": true, "isNewDevice": true,
  "license": {email, status, maxDevices, activatedAt},
  "device": {deviceId, firstSeenAt} }
```

### 2.4 Static

| Path | Purpose |
|---|---|
| `GET /*` | Serve `public/` (taskpane UI) |
| `GET /assets/*` | Serve `assets/` |
| `GET /favicon.ico` | Return `assets/icon-32.png` |
| `GET /manifest.xml` | Office Add-in manifest (Office cần fetch tại root) |

---

## 3. WebSocket — `/ws`

### 3.1 Connection
- **Path**: `ws[s]://&lt;host&gt;:&lt;port&gt;/ws`
- **Auth**: License + Device qua query string `?k=&lt;licenseKey&gt;&amp;d=&lt;deviceId&gt;`
  (browser WebSocket API không gửi custom headers)
- **verifyClient** (`server.js:202-214`): nếu `LICENSE_REQUIRED=true` mà thiếu `?k` hoặc `?d` → reject HTTP 401

### 3.2 Heartbeat
- Server ping/pong mỗi **15s** (`server.js:248-254`)
- `ws.isAlive=false` → `ws.terminate()`
- Client reconnect sau **2.5s** khi `onclose`/`onerror` (`ws-client.js:113-119`)

### 3.3 Message Protocol

**Server → Client (op request)**:
```json
{ "type": "op", "id": "<uuid>", "op": { "kind": "findReplace", ... } }
```

**Client → Server (kết quả)**:
```json
{ "type": "result", "id": "<uuid>", "ok": true, "result": ... }
{ "type": "result", "id": "<uuid>", "ok": false, "error": "...", "detail": ... }
{ "type": "hello", "kind": "word", "info": "Microsoft Word Taskpane" }
```

---

## 4. MCP Tools — 18 tools (stdio)

Server name: `word-bridge-mcp`. Dispatcher: `mcp-server.js:39-95` (strip prefix `word_` để lấy `kind`).

| Kind | Mô tả |
|---|---|
| `ping` | Kiểm tra kết nối 2 chiều |
| `getText` | Đọc toàn bộ văn bản doc (limit mặc định 8000) |
| `getParagraphs` | List đoạn văn có style + text tóm tắt |
| `getSelection` | Đọc text user đang bôi đen |
| `selectRange` | Bôi đen range theo `anchor` / `paragraphIndex` |
| `findReplace` | Tìm & thay toàn bộ doc (optionally Track Changes) |
| `insertAfterText` | Chèn text ngay sau anchor |
| `deleteText` | Tìm & xoá text |
| `setParagraphStyle` | Gán style cho đoạn văn (Heading, Normal, Quote...) |
| `insertComment` | Tạo Word Comment ở lề phải |
| `insertOoxml` | Chèn OOXML fragment |
| `getTrackedChanges` | List tracked changes (3s timeout) |
| `setTrackChanges` | Bật/tắt Track Changes |
| `getMetadata` | Đọc Title/Author/Subject |
| `getComments` | List margin comments |
| `acceptComment` | Apply suggestion: replace text + xoá comment |
| `deleteComment` | Xoá comment theo anchor |
| `replyComment` | Reply threaded vào comment |

### 4.1 ⚠️ Bất cập giữa `word-tool-configs/` (client) và `tools.js` (server)

| Configs (client, 17) | tools[] (server, 18) | Note |
|---|---|---|
| `word_read_selection` | `word_getSelection` | ✓ match |
| `word_read_document` | `word_getText` | ✓ |
| `word_read_paragraphs` | `word_getParagraphs` | ✓ |
| `word_search_text` | ❌ KHÔNG có trong tools[] | Client có, server reject |
| `word_insert_text` | `word_insertAfterText` | name khác |
| `word_insert_markdown` | ❌ KHÔNG có | Client có, server reject |
| `word_insert_comment` | `word_insertComment` | ✓ |
| `word_accept_comment` | `word_acceptComment` | ✓ |
| `word_resolve_comment` | ❌ KHÔNG có | Client có, server reject |
| `word_insert_with_track_changes` | ❌ KHÔNG có | Client có, server reject |
| `word_get_tracked_changes` | `word_getTrackedChanges` | ✓ |
| `word_find_replace` | `word_findReplace` | ✓ |
| `word_delete_text` | `word_deleteText` | ✓ |
| `word_set_paragraph_style` | `word_setParagraphStyle` | ✓ |
| `word_select_range` | `word_selectRange` | ✓ |
| `word_insert_ooxml` | `word_insertOoxml` | ✓ |
| `word_get_text_analytics` | ❌ KHÔNG có | Client có, server reject |

**`CONFIG_TO_MCP_TOOLS` mapping** ở `tools.js:385-404` tồn tại nhưng `dispatchViaConfig()` chỉ lookup rồi fallback `clientManager.dispatchOp` — **không auto-load schema từ configs**. Khi client gọi các kind không có trong `tools[]`, `ws-client.js:202-204` throw `Công cụ không được hỗ trợ: &lt;kind&gt;`.

**Cần fix**: auto-generate `tools[]` từ `word-tool-configs` (single source of truth) HOẶC thêm 5 kind còn thiếu vào `tools.js` (`searchText`, `insertMarkdown`, `resolveComment`, `insertWithTrackChanges`, `getTextAnalytics`).

---

## 5. AI Proxy Config

### 5.1 Endpoint
- **Mặc định**: `https://aiapi.2tocom.space/v1/chat/completions` (gateway riêng)
- **Anthropic detect**: URL chứa `anthropic` / `/messages` → dùng headers Anthropic
- Ngược lại → OpenAI-compatible (Bearer token)

### 5.2 Model
- **Mặc định**: `MiniMax-M3` (custom gateway)
- Client **không override** được — server luôn dùng config của mình

### 5.3 Effort → thinking mapping

| Effort | Anthropic `thinking.budget_tokens` | OpenAI `reasoning_effort` |
|---|---|---|
| `off` | (không gửi) | (không gửi) |
| `low` | 1024 | `low` |
| `medium` | 4096 | `medium` |
| `high` | 16384 | `high` |
| `xhigh` | 32768 | `xhigh` |

Constraint: `max_tokens &gt; budget_tokens` (Anthropic). Nếu vi phạm → bump `max_tokens = budget_tokens + 1024`.

### 5.4 Headers
- **Anthropic**: `x-api-key` + `anthropic-version: 2023-06-01` + `anthropic-dangerous-direct-browser-access: true`
- **OpenAI**: `Authorization: Bearer &lt;key&gt;`
- API key lưu ở env `AI_API_KEY` (fallback `"keykeykeykeykeykeykeykey"` cho dev)

### 5.5 Retry Logic
**Chỉ 1 retry duy nhất**:
- Trigger: upstream 4xx + `effort === "xhigh"` + provider **không phải** Anthropic
- Action: rebuild với `effort: "high"` rồi fetch lại
- Nếu vẫn fail → trả 4xx với `note: "Đã retry với effort=high nhưng vẫn fail."`

---

## 6. SSE Forwarding

Server forward 4 event types từ upstream AI về client:

| Event | Trigger |
|---|---|
| `delta` | `delta.text` (Anthropic) HOẶC `choice.delta.content` (OpenAI) |
| `thinking` | `delta.thinking` (Anthropic) HOẶC `choice.delta.reasoning_content` (DeepSeek R1) |
| `done` | Stream kết thúc |
| `error` | Lỗi giữa chừng stream |

Response headers:
```
Content-Type: text/event-stream; charset=utf-8
Cache-Control: no-cache, no-transform
Connection: keep-alive
X-Accel-Buffering: no
```

---

## 7. License System

### 7.1 Database Schema (3 tables + 2 RPC)

**`public.licenses`**: `id, license_key, email, status, max_devices (1-10), notes, activated_at, created_at, updated_at`

**`public.license_devices`**: `id, license_id (FK CASCADE), device_id (32 hex), device_name, ip_address, user_agent, first_seen_at, last_seen_at, revoked_at`
- Unique partial index: `(license_id, device_id) WHERE revoked_at IS NULL`

**`public.license_events`**: log `activate/heartbeat/revoke/reject/info`

**RPC `enroll_device(...)`**: atomic FOR UPDATE, touch `last_seen_at` hoặc insert mới + log event. Trả `LICENSE_NOT_FOUND` / `LICENSE_INACTIVE` / `DEVICE_LIMIT_EXCEEDED` / `ok`.

**RPC `validate_device(...)`**: check + touch `last_seen_at`. Trả `LICENSE_NOT_FOUND` / `LICENSE_INACTIVE` / `DEVICE_NOT_ENROLLED` / `ok`.

RLS bật cho 3 bảng (service_role bypass, anon key blocked).

### 7.2 Activate → Heartbeat → Devices Flow

```
Client → POST /api/license/activate {email, licenseKey, deviceId}
       ← {ok, isNewDevice, license, device}

Mỗi 5-10 phút:
Client → POST /api/license/heartbeat
       ← {ok: true, ttl: 3600}

Khi cần list devices:
Client → GET /api/license/devices
       ← {license, devices: [{id, deviceId, deviceName, ip, userAgent, firstSeenAt, lastSeenAt, revokedAt, isCurrent}]}

Khi muốn thu hồi 1 device:
Client → DELETE /api/license/device/{deviceId}
       ← {ok, revoked: true} hoặc 409 nếu là device cuối
```

---

## 8. Config Defaults (env vars)

| Biến | Mặc định | Note |
|---|---|---|
| `PORT` | `3011` | |
| `HOST` | `0.0.0.0` | |
| `USE_HTTPS` | `false` | Cần `office-addin-dev-certs` nếu true |
| `PUBLIC_URL` | `https://word-bridge.2tocom.space` | |
| `LICENSE_REQUIRED` | `true` | `false` = dev mode (no auth) |
| `AI_API_ENDPOINT` | `https://aiapi.2tocom.space/v1/chat/completions` | |
| `AI_API_KEY` | `"keykeykeykeykeykeykeykey"` (dev) | |
| `AI_MODEL` | `MiniMax-M3` | |
| `AI_DEFAULT_EFFORT` | `medium` | |
| `AI_MAX_TOKENS` | `4096` | |
| `AI_TEMPERATURE` | `0.7` | |
| `LICENSE_ACTIVATE_RATE_PER_MIN` | `5` | |
| `LICENSE_HEARTBEAT_TTL_SEC` | `3600` | |
| `SUPABASE_URL` | (required nếu license) | |
| `SUPABASE_SERVICE_KEY` | (required nếu license) | |
| `CORS_ORIGINS` | `localhost:3001, *.officeapps.live.com` | Wildcard `*` OK |

---

## 9. CORS + Static

CORS whitelist qua `config.corsOrigins` (wildcard `*` supported). `app.set("trust proxy", 1)` để respect `X-Forwarded-For` từ Cloudflare.

Static serving: `express.static(publicDir, {etag: false, maxAge: 0})` — no cache (dev mode).

---

## 10. Boot URLs (`server.js:265-276`)

```
🚀 AI WORD ADD-IN & MCP AGENT BRIDGE SERVER
• Web / Taskpane URL : https://localhost:3011/index.html
• WebSocket Bridge   : wss://localhost:3011/ws
• AI Chat Proxy      : https://localhost:3011/api/chat
• Tool Catalog       : https://localhost:3011/tools
• Status Check       : https://localhost:3011/api/status
```

---

## 11. Files Map

| File | LOC | Chức năng |
|---|---:|---|
| `server.js` | 276 | Entry: Express + CORS + routes mount + WebSocketServer + heartbeat |
| `ai-proxy.js` | 339 | SSE forward Anthropic/OpenAI, effort mapping, retry logic |
| `mcp-server.js` | 110 | MCP stdio server (ListTools + CallTool) |
| `tools.js` | 421 | 18 tool definitions + `WordClientManager` (WebSocket dispatch) |
| `config.js` | 134 | Env config, validate, log summary |
| `license/routes.js` | 233 | Express router + 6 endpoints + `requireLicense` middleware |
| `license/license-service.js` | 268 | Business logic: activate / validate / heartbeat / revoke |
| `license/schema.sql` | - | DB schema + 2 RPC functions |
| `license/device-fingerprint.js` | 36 | sha256(IP+UA+ClientId) |
| `license/supabase-client.js` | 33 | Singleton Supabase client (service_role) |
| `license/middleware.js` | 6 | Re-export `requireLicense` |

**Tổng server**: ~1856 LOC.

---

## 12. Known Issues / TODO

1. **`word-tool-configs` không sync với `tools[]`** — 5 configs (searchText, insertMarkdown, resolveComment, insertWithTrackChanges, getTextAnalytics) không có trong `tools[]`. Client gọi các kind này sẽ bị server reject.
2. **`device-fingerprint.js` không dùng** trong `license-service.js` — code có sẵn nhưng bypass.
3. **`ping`/`getMetadata` thiếu dispatcher** ở `ws-client.js` (chỉ khai báo schema server, client chưa wire) — nhưng `ping` có fallback.
4. **Schema validate body chưa chặt** — `/api/chat` và `/op` không dùng AJV/zod; client gửi gì server dùng đó (đã strip `endpoint/apiKey/model` ở `ai-proxy.js:140-143`).
5. **Rate limit chỉ ở `/activate`** — `/api/chat`, `/op` không giới hạn → dễ abuse nếu expose public.
