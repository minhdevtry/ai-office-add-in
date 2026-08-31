# OfficeCLI — Research Note cho Ori AI

> Tài liệu tóm tắt repo [`iOfficeAI/OfficeCLI`](https://github.com/iOfficeAI/OfficeCLI) (29.6k stars, 6,079 commits, Apache 2.0).
> Mục đích: rút ra patterns có thể áp dụng cho Ori AI Word Add-in.
> Đọc xong tài liệu này có thể quyết định hướng tiếp theo mà không cần đào sâu thêm.
> Cập nhật lần cuối: 2026-09-01.

---

## 1. Tổng quan

**OfficeCLI** là một binary C# .NET duy nhất (`officecli`) đọc, sửa, tự động hoá file `.docx` / `.xlsx` / `.pptx` mà **không cần Microsoft Office cài sẵn** — thao tác trực tiếp lên OOXML (Open XML) bằng OpenXml SDK.

**Khẩu hiệu:** *"World's first and the best Office suite designed for AI agents."*

**Đặc điểm:**
- Binary self-contained, multi-platform (macOS/Linux/Windows ARM64+x64), cài qua curl / Homebrew / Scoop / npm wrapper
- Không tích hợp LLM — chỉ là tool deterministic, **AI agents (Claude Code, Cursor, Windsurf, Copilot, Codex) dùng nó qua MCP server**
- Có SDK Node.js + Python mỏng gọi qua named-pipe IPC vào một "resident" process

**Điểm khác biệt với Ori AI:**
- OfficeCLI = CLI cho file Office (không cần Office)
- Ori AI = Office Add-in (cần Office mở + add-in sideload)
- Cả hai đều phục vụ AI agent thao tác tài liệu — OfficeCLI từ terminal/CI, Ori AI từ sidebar Word

---

## 2. Tại sao nghiên cứu OfficeCLI?

OfficeCLI là codebase mature nhất hiện tại về "AI agent + Office documents". 4 patterns đáng học:

1. **3-Layer architecture** rất rõ (Read / DOM / Raw XML) — agent không phải đoán layer nào dùng
2. **Self-healing errors** với structured codes + Levenshtein auto-correct — DX cho agent cực tốt
3. **JSON Schema contracts** làm source of truth cho mọi prop/alias — dev workflow chặt chẽ
4. **Resident mode** (in-process document cache) + per-file singleton — giải quyết vấn đề state mà Ori AI đang stateless

---

## 3. Kiến trúc 3-Layer

`src/officecli/Core/IDocumentHandler.cs:53-114` định nghĩa interface `IDocumentHandler` chia rõ 3 layer:

| Layer | Mục đích | Methods | Idempotent? |
|---|---|---|---|
| **L1 Read** | Semantic views | `ViewAsText`, `ViewAsAnnotated`, `ViewAsOutline`, `ViewAsStats`, `ViewAsStatsJson`, `ViewAsIssues` | ✅ |
| **L2 DOM** | Element operations | `Get`, `Query`, `Set`, `Add`, `Remove`, `Move`, `CopyFrom` | ❌ |
| **L3 Raw XML** | XPath fallback | `Raw`, `RawSet`, `AddPart`, `TryExtractBinary` | ❌ |

**Quy tắc vàng trong SKILL.md:** *"L1 → L2 → L3. Always start with L1 (view) to understand the document, then L2 for mutations, then L3 only when L2 isn't enough."*

**Áp dụng cho Ori AI:** MCP tools nên có layer tương tự:
- L1 read: `view_document(mode)`, `get_paragraph(path)`, `find_issues()`
- L2 mutate: `set_paragraph`, `add_paragraph`, `remove_paragraph`
- L3 raw: `read_xml(part)`, `replace_xml(part, xpath, fragment)`

Agent sẽ thấy rõ "cái này không thuộc L2 → fall through L3" — giống quy tắc trong skill.

---

## 4. Path DSL `/slide[1]/shape[2]`

Path system riêng của OfficeCLI (không phải XPath), parser per-handler:

**Cú pháp:** `/segment1[N]/segment2[@attr=value]/...`
- **1-based** indexing ở user-facing (giống XPath)
- **0-based** ở array code — `PathIndex.ToArrayIndex` / `FromArrayIndex` tại `src/officecli/Core/PathIndex.cs:23-30` làm ranh giới duy nhất
- **Local names** (không namespace): `p` cho paragraph, `r` cho run, `tbl` cho table, `tr` cho row, `tc` cho cell
- **Stable IDs** ưu tiên khi có: `/body/p[@paraId=ABC]` thay vì `/body/p[3]`
- **Aliases**: `p` ≡ `paragraph`, `tr` ≡ `row`, ...

**Word parser:** `src/officecli/Handlers/Word/WordHandler.Navigation.cs:714+` (`ParsePath`)
- Bắt đầu bằng `/`, reject trailing `/`, reject `//`
- Selector operators: `=`, `!=`, `~=`, `>=`, `<=`
- Combinator `>` (child), `:contains`, `:empty`, `:has`, `:no-alt`
- Boolean `and` / `or` cho compound selectors

**Áp dụng cho Ori AI:** Word document path theo Word XML local names. 1-based + ưu tiên `@paraId` khi có. Centralize `<-> 0-based` boundary. Agent sẽ viết được path; predictable; Office UI cũng dùng 1-based khi user click "Paragraph 3".

---

## 5. MCP — "1 thin shell tool" pattern

Đây là pattern **gây tranh cãi nhất** trong codebase. OfficeCLI chỉ expose **DUY NHẤT 1 tool** tên `officecli` với 1 param `command` (string CLI-like):

```
Tool: officecli
Param: { command: "docx view foo.docx text" }  → result
       { command: "docx set foo.docx /body/p[3] --prop bold=true" }  → result
```

Tại `src/officecli/McpServer.cs:208-260` có comment giải thích lý do (paraphrase):

> *"No per-command marshalling here means no argument can be silently dropped (every CLI flag works for free), and the model writes exactly what the skills' CLI examples show."*

**Cách hoạt động:**
- Tool nhận `command` → `Tokenize` (quote-aware) → strip leading `officecli` → `RootCommand.Parse` qua **cùng `System.CommandLine` root mà CLI dùng** → exec, capture stdout+stderr, return content block
- JSON-RPC 2.0 thuần, hand-rolled `Utf8JsonWriter` (no reflection, PublishTrimmed-safe)
- Methods: `initialize`, `notifications/initialized`, `tools/list`, `tools/call`, `ping`
- `screenshot` handle riêng: chạy CLI, đọc PNG, trả `image` content block base64

**So sánh với Ori AI hiện tại:**
- Ori AI `server/mcp-server.js` expose **18 tools**, mỗi tool 1 verb riêng (`word_getText`, `word_findReplace`, ...)
- Trade-off:
  - **1 tool (OfficeCLI)**: thêm verb = thêm CLI sub-command, MCP schema không đổi, model tự nhớ cú pháp
  - **18 tools (Ori AI)**: model ít phải nhớ verb, DX tốt hơn cho agent, nhưng mỗi tool tốn token context, thêm verb = edit MCP schema
- **Có thể hybrid**: 1 thin shell tool + wrapper ngắn cho top-3 verb thường dùng

---

## 6. SDK + Named-Pipe IPC

**Cấu trúc `sdk/`:**
- `sdk/node/index.js` (~600 dòng) + `index.d.ts` (typed)
- `sdk/python/officecli.py` (mirror logic)
- Cả 2 chỉ là **thin shell forward commands** qua named-pipe vào resident process. Không tự xử lý file.

**API surface Node:**

```ts
create(filePath, args?, options?) → Document
open(filePath, options?) → Document

class Document {
  path: string
  send(item, asJson?, timeoutMs?) → Result
  batch(items, options?) → Result
  alive(timeoutMs?) → boolean
  close()
  [Symbol.asyncDispose]()  // Node 24+
}

install()  // chạy official installer
pipePaths(filePath) → [main, ping]  // debug helper
```

Functional API, **Promise-based**, async/await native. Class `OfficeCliError extends Error` (code: number) chỉ raise khi **transport failure** (không phải business error; business error nằm trong envelope `success:false`).

**Named-pipe IPC:**
- Tên pipe: `officecli-<SHA256(canonical_fullpath)[:16] UPPERCASE>` — `sdk/node/index.js:96-110`
- Unix: `$TMPDIR/CoreFxPipe_<name>` + suffix `-ping`
- Windows: `\\.\pipe\<name>` + `-ping`
- Framing: 1 request line + 1 response line, UTF-8, `\n`-terminated, **1 connection = 1 command**
- **Hai pipes song song**: main + `-ping` (liveness probe, response ngay cả khi main busy)
- Reply có UTF-8 BOM + trailing `\r` — wrapper strip cả hai

**Retry policy:**
- Connect timeout 30s, max 3 retries với backoff `50*(n+1)ms`
- **Retries chỉ re-attempt connect (trước khi command chạy) → re-send an toàn cho mutation**
- **Empty reply = crash mid-serve**: raise KHÔNG retry (vì command có thể đã apply)

**Resident mode** (`src/officecli/Core/ResidentServer.cs:1-100`):
- Process con spawn bởi CLI đầu tiên chạm file
- Giữ `WordprocessingDocument`/`SpreadsheetDocument`/`PresentationDocument` mở
- **Autosave** 2-10s adaptive (tắt qua `OFFICECLI_RESIDENT_FLUSH=off`)
- **Idle shutdown** 12 phút interactive / 60s sau `create`
- **Per-file singleton** (`CommandBuilder.cs:126-150`): probe `ResidentClient.TryConnect` trước khi spawn, chống 2 resident mở cùng file

**Quote SDK comment (lines 9-15):**

> *"Bootstrap vs hot-path are deliberately separated: `create`/`open` spawn a process; `send`/`batch` is a pipe round-trip, never a spawn. The hot path."*

**Áp dụng cho Ori AI (1 phần):**
- **Server-side document state**: `Map<filePath, DocumentState>` trong Node server, mỗi WS session bind vào 1 state
- **Bootstrap vs hot-path**: lần đầu mở file → server mở Office doc + cache; op tiếp theo chỉ touch cache, save lazy
- **Resident-per-file singleton**: chống race khi user edit trong Word UI + agent cùng mutate
- **2 channel IPC**: WS main + WS ping (heartbeat) — pattern giống main+ping pipe
- **Lưu ý**: WS qua mạng khác named-pipe cục bộ, retry phải cẩn thận hơn — cần idempotency-key cho non-idempotent op

---

## 7. Error Handling + Self-Healing

### 7.1 Structured CliException

`src/officecli/Core/CliException.cs:1-27`:

```csharp
class CliException : Exception {
  string Code          // not_found, invalid_value, ...
  string? Suggestion   // "Did you mean: bold?"
  string? Help         // "Run: help word paragraph"
  string[]? ValidValues  // ["left","center","right"]
}
```

**Codes thường gặp:**
`not_found`, `invalid_value`, `invalid_path`, `invalid_render`, `file_not_found`, `file_required`, `unsupported_type`, `unsupported_property`, `internal_error`, `auto_corrected`, `zero_matches`, `text_overflow`, `advisory`, `unrecognized_latex_command`, `missing_prop_flag`

### 7.2 Self-correction (Damerau-Levenshtein)

`src/officecli/Core/EditDistance.cs:17-39` — implementation duy nhất:

> *"The most common real-world typo class (`blod`→`bold`, `Salray`→`Salary`) scores 1 instead of Levenshtein's 2 and stays inside tight suggestion thresholds."*

- Damerau restricted (swap kề = 1) thay vì Levenshtein thường
- Threshold: `max(2, len/3)`
- **Substring preference trước Levenshtein** — `paragrah` → `paragraph` (catch trước khi tính distance)
- **Tie-break uniqueness**: chỉ suggest khi "best" unique (ambiguous → không suggest, an toàn hơn)
- **Scope filter** theo format (excel/word/pptx) — loại prop chỉ-hợp-format-khá
- **Dotted-segment matching**: so sánh cả `bold` lẫn `font.bold`

### 7.3 TrackingPropertyDictionary — "handler-as-truth"

`src/officecli/Core/TrackingPropertyDictionary.cs:1-160`:

> Pattern thay thế "schema pre-filter" cũ. Schema không còn runtime gate, **handler là source of truth**.

- Detect "user supplied `--prop X` but handler never read X" qua custom `IEqualityComparer`
- Override `GetEnumerator` + re-declare `IEnumerable<KVP>` để LINQ `.Where()` cũng bị track
- Mọi prop không dùng → warning thay vì silently drop

### 7.4 JSON Envelope chuẩn

`src/officecli/Core/OutputFormatter.cs:103-200`:

```json
// Success
{ "success": true, "data": <node>|<text>, "message": "Updated /body/p[3]: bold=true",
  "warnings": [...], "matched": N }

// Error
{ "success": false,
  "error": { "code": "not_found", "message": "...",
             "suggestion": "...", "help": "...", "validValues": ["..."] } }
```

- `success:false` = business verdict (apply thất bại)
- Exit code chỉ dành cho transport error
- MCP mapping: `exit 1`/`exit 2 + no stdout` = `isError:true`; `exit 2 + stdout` = `isError:false` (apply with caveats)

### 7.5 Warning context (thread-static)

`OutputFormatter.cs:82-101` — `WarningContext.Begin/Add/End`:
- Handler sâu không với tới biến local của command layer
- Thread-static accumulator; envelope wrapper gom vào `warnings[]`
- Pattern: deep helper muốn warn (vd "text overflow detected") không cần truyền outparam qua 5 lớp

### 7.6 Warning kinds

```typescript
type WarningKind =
  | "unknown_key"        // prop name không tồn tại
  | "value_no_match"     // value không thuộc validValues
  | "non_numeric"        // expect number, get string
  | "auto_corrected"     // Levenshtein đã sửa
  | "zero_matches"       // find/replace không match
  | "text_overflow"      // text vượt container
  | "advisory"           // thông tin, không fail
  | "unrecognized_latex_command"
  | "missing_prop_flag";  // --prop flag missing
```

**So sánh với Ori AI hiện tại:**
- Đang throw raw exception → `err.message || String(err)` qua WebSocket
- Model phải parse text → đoán → retry
- OfficeCLI cải thiện 5 chiều:
  1. **Discriminated error code** — model phân biệt `not_found` (retry path khác) vs `invalid_value` (đổi value)
  2. **`suggestion` + `validValues` machine-readable** — model tự sửa được
  3. **Self-correction trong tool** — `set ... --prop colot=red` tự thành `color=red`
  4. **Non-fatal warnings** — `zero_matches` thành `warnings[]` không fail
  5. **Idempotency hint** — empty reply = "không retry"

---

## 8. JSON Schema Contracts

### 8.1 `schemas/help/` directory

```
schemas/help/
├── _schema.json              # JSON Schema (draft 2020-12) meta-schema
├── docx/<element>.json       # Word: paragraph, run, table, comment, ...
├── xlsx/<element>.json       # Excel: cell, row, sheet, chart, ...
├── pptx/<element>.json       # PowerPoint: slide, shape, picture, ...
└── _shared/<element>.json    # Shareable giữa formats
```

### 8.2 Schema file structure

`schemas/help/_schema.json` defines meta-schema. Required fields:
- `format`, `element`, `operations`, `properties`
- Optional: `paths`, `addressing`, `children`, `parts`, `addParent`, `aliases`, `note`, `description`, `container`, `extends`, `elementAliases`

`$defs/property` chi tiết:
- `type`, `description`, `values`, `aliases`
- `add`/`set`/`get` (operations supported)
- `examples`, `readback`, `enforcement` (`"strict"` hoặc `"report"`)

### 8.3 Validate input ở đâu?

**Hai tầng validation:**
1. **CLI parse** — `System.CommandLine` (built-in): required arg, type convert, value validate qua custom validator
2. **Handler-as-truth** — `TrackingPropertyDictionary` (đã nói ở §7.3) thay thế "schema pre-filter" cũ

**Runtime schema loader:** `src/officecli/Help/SchemaHelpLoader.cs:1-200`:
- Load embedded JSON
- Normalize format alias (`word→docx`, `excel→xlsx`, `ppt/powerpoint→pptx`)
- Resolve `elementAliases` (so `help docx p` ≡ `help docx paragraph`)
- `ClosestMatch` cho unknown element (line 189)

### 8.4 CI contract test

`schemas/README.md` ghi:

> *"Any PR that changes Add/Set/Get for an element MUST update the matching schema file in the same PR. CI contract tests will fail otherwise."*

Test thực tế (`tests/OfficeCli.NumericFit.Tests/Program.cs`, ~600 dòng):
- Integration test, không có unit test mock
- Tạo fixture `.xlsx` trong temp, mở qua `ExcelHandler`, assert `ViewAsIssues()` trả đúng path/severity/subtype/suggestion
- `enforcement: "strict"` → break CI; `"report"` → log only
- Không có snapshot test truyền thống — chỉ issue-list assertions

### 8.5 Help system

```bash
officecli help <format> <element>              # schema cho element
officecli help <format> <verb> <element>       # schema cho verb+element
officecli help <format> <element> --json       # machine-readable
```

**Áp dụng cho Ori AI:**
- Tạo `server/schemas/<element>.json` (canonical prop names + aliases + examples)
- Tool `help(element, verb?)` trả về
- CI: any PR thay đổi tool behavior phải update schema; otherwise fail
- Agent không phải đoán prop name; dev thêm prop mới phải update schema

---

## 9. SKILL.md + Auto-Install

### 9.1 SKILL.md format (chuẩn `agentskills.io`)

File `/tmp/officecli-research/SKILL.md` (~25 KB). Frontmatter YAML (`name`, `description`) rồi Markdown. Vai trò: **tài liệu routing cho AI agent**, không phải spec nội bộ.

Các phần chính:
- **Install** — curl install.sh / install.ps1
- **Strategy** — L1 → L2 → L3, luôn `--json` khi structured, **nhắc load skill chuyên dụng trước** (gate)
- **Help System (IMPORTANT)** — quy tắc "không bao giờ đoán prop name, chạy `help`"
- **Performance: Resident Mode** — auto-start 60s, explicit `open`/`close` 12 min, opt-out `OFFICECLI_NO_AUTO_RESIDENT=1`
- **Quick Start** — PPT/Word/Excel
- **L1/L2/L3 sections** — command reference đầy đủ
- **Stable ID Addressing** — `/slide[1]/shape[@id=N]`, `/body/p[@paraId=X]`
- **Query selectors** — CSS-like với `=`, `!=`, `~=`, `>=`, `<=`, `:contains`, `:empty`, `:has`
- **Watch & Interactive Selection** — live HTML preview, browser click chọn shapes, `get selected`
- **L2: DOM Operations** — `set`, `find`/`replace`, `add` (`--from` clone, `--after`/`--before`/`--index`)
- **L3: Raw XML** — `raw`, `raw-set` (actions: append/prepend/insertbefore/insertafter/replace/remove/setattr)
- **Common Pitfalls** (~13 lỗi thường gặp)
- **Specialized Skills** — routing bảng: `word` vs `academic-paper`, `pptx` vs `pitch-deck` vs `morph-ppt` vs `morph-ppt-3d`, ...
- **Notes** — paths 1-based, `--index` 0-based (trừ `add --type row|col` xlsx)

### 9.2 `skills/` directory — 11 skills

- `officecli` — base skill (umbrella)
- `officecli-docx`, `officecli-xlsx`, `officecli-pptx` — per-format build guide
- `officecli-pitch-deck` — fundraising-only (seed, Series A-C, SAFE, convertible)
- `officecli-academic-paper` — journal/conference (APA/Chicago/IEEE/MLA)
- `officecli-data-dashboard` — CSV/tabular → KPI dashboard
- `officecli-financial-model` — financial projections
- `officecli-word-form` — fillable forms, content controls
- `morph-ppt` — cross-slide Morph animation
- `morph-ppt-3d` — 3D Morph (GLB models, camera)

Mỗi skill có SKILL.md + (thường) `reference/*.md` + style/template binary.

### 9.3 `load_skill` lazy reference

`McpServer.cs:541-590` — `McpHelpStrategy` + `BuildSkillTriggerSummary`:
- Inject trực tiếp 1-line trigger vào MCP tool description
- Agent tự quyết định `load_skill <name>` nào
- `load_skill` đọc embedded resource, append manifest reference files, return text
- Binary assets (template, theme) cài qua `officecli install` riêng

### 9.4 Auto-install

`src/officecli/Core/SkillInstaller.cs:11-25` định nghĩa bảng 13 agent targets (claude, copilot, codex, cursor, pi, windsurf, ...):
- Detect agent qua sự tồn tại folder config (không registry trung tâm)
- `InstallBaseToAll` (line 415-440) quét `Home/.claude`, `Home/.copilot` … copy `SKILL.md`
- `McpInstaller.cs:80-130` install binary + MCP config cho 4 client (claude, cursor, vscode, lmstudio)

**Update mechanism:** `McpServer.cs:73-110` — background check mỗi giờ, debounce 24h qua `~/.officecli/config.json`. Download qua signed subprocess (stdio redirected) để không corrupt MCP stdout.

---

## 10. Các pattern khác đáng học

### 10.1 Idempotency hint cho non-idempotent ops
- SDK Node: empty reply = "command có thể đã apply, KHÔNG retry"
- Quan trọng khi IPC qua mạng (WebSocket reconnect)
- Cần request-id + server cache response 30s

### 10.2 Watch mode: live preview + browser selection
- `Core/Watch/WatchServer.cs` (SSE relay)
- `Mark`/`Unmark`/`Marks`/`Goto` qua pipe IPC
- Browser click chọn shapes → CLI `get selected` trả paths
- Cho Ori AI: server có thể poll `Office.context.document.getSelection()` qua WS heartbeat

### 10.3 Audit trail: `docProps/custom.xml` stamp khi mutate
- `WordHandler.Modified` flag → `Dispose` stamp `OfficeCLI metadata`
- Pure Get sessions không touch
- Cho Ori AI: ghi audit log vào custom XML part (author, timestamp, op, agent ID) — chỉ khi có mutation

### 10.4 Per-handler warning context
- Thread-static accumulator (AsyncLocal trong JS)
- Deep helper không cần truyền outparam qua 5 lớp
- Top-level response builder flush warnings vào envelope

### 10.5 Element alias layer
- `elementAliases` trong schema: `paragraph` ≡ `p`
- `propAliases` trong schema: `align` ≡ `alignment` ≡ `halign`
- Một resolver trung tâm map segment → DOM element

---

## 11. Tổng hợp patterns áp dụng cho Ori AI

Xếp theo **impact** (cao → thấp) và **effort** (thấp → cao):

| Pri | Pattern | Effort | Impact | Phụ thuộc |
|---|---|---|---|---|
| **P0** | **P3 CliException structured** (code/suggestion/validValues) | 1 ngày | Fix error UX ngay | — |
| **P0** | **P6 JSON envelope chuẩn** `{success, data, warnings}` | 0.5 ngày | Mọi tool cùng schema | P3 |
| **P0** | **P4 Levenshtein self-correct** prop name | 1 ngày | DX cho agent cải thiện mạnh | KnownProps registry |
| P1 | P1 Server-side document state (resident pattern) | 3-5 ngày | Latency + consistency | — |
| P1 | P5 3-layer verbs exposed qua MCP | 1-2 ngày | Agent biết layer nào dùng | P6 |
| P1 | P7 Help system + JSON Schema contract | 2-3 ngày | Dev workflow + agent path | P6 |
| P2 | P2 MCP "1 thin shell tool" (thay 18 tools) | 1 ngày | Token savings | — |
| P2 | P8 Path DSL 1-based với stable IDs | 1 ngày | Agent path writing | P5 |
| P2 | P10 Lazy skill load qua `load_skill` | 1 ngày | Context savings | P7 |
| P3 | P9 Idempotency hint cho non-idempotent ops | 1 ngày | An toàn khi WS reconnect | — |
| P3 | P15 Audit trail vào custom XML part | 1 ngày | Nice-to-have | — |
| P3 | P11 Aliases (elementAliases + propAliases) | 0.5 ngày | Convenience | P7 |
| P3 | P12 AsyncLocal warning | 0.5 ngày | Dev convenience | P6 |
| P3 | P13 Watch mode (poll selection) | 1-2 ngày | Nice-to-have | — |

**Khuyến nghị thứ tự áp dụng:**

### Giai đoạn 1 (1 tuần) — Quick wins
- P0 × 3: structured error + JSON envelope + Levenshtein self-correct
- Ưu điểm: không phụ thuộc nhau, dễ review từng commit, cải thiện DX model + UX user ngay lập tức

### Giai đoạn 2 (2 tuần) — Foundation
- P1 Server-side document state (resident pattern)
- P5 3-layer verbs + P6 envelope (đã có sẵn từ P0)
- P7 Help system + JSON Schema contract
- P2 MCP "1 thin shell tool" (có thể làm song song)

### Giai đoạn 3 (1-2 tuần) — Polish
- P8 Path DSL với stable IDs
- P10 Lazy skill load
- P9 Idempotency hint
- P15 Audit trail

---

## 12. So sánh tổng quan Ori AI vs OfficeCLI

| Khía cạnh | Ori AI hiện tại | OfficeCLI | Học được |
|---|---|---|---|
| Runtime | Office Add-in (cần Word) | CLI standalone (không cần Office) | Khác nhau — giữ nguyên |
| Document access | Office.js API (Word live) | OpenXml SDK (file trực tiếp) | Khác nhau — giữ nguyên |
| MCP tools | 18 tools riêng | 1 thin shell tool | Có thể học (P2) |
| Error format | `err.message` raw text | `{code, message, suggestion, validValues}` | **Học ngay (P0)** |
| Prop correction | Không có | Damerau-Levenshtein auto-fix | **Học ngay (P0)** |
| JSON envelope | Ad-hoc per tool | `{success, data, warnings}` chuẩn | **Học ngay (P0)** |
| State | Stateless (mỗi op đi qua WS) | Resident mode per-file singleton | Học sau (P1) |
| Path system | Implicit (anchor text) | DSL 1-based với stable IDs | Học sau (P2) |
| Schema | `word-tool-configs/` 17 tools | Per-element JSON Schema | Học sau (P1) |
| Help system | Không có | `help <element>` | Học sau (P1) |
| Skills | 8 skills (skills.json) | 11 skills + 3 specialized | Tham khảo format |
| License | License + Supabase RPC | Không có (open source) | Khác — giữ nguyên |
| Stream proxy | 3-layer (StreamProxy/CORS/Direct) | Không có | Khác — giữ nguyên |
| SSE | 4 event types (delta/thinking/done/error) | Không có (1-line response) | Khác — giữ nguyên |

---

## 13. File references chính

| File | Vai trò |
|---|---|
| [`src/officecli/Core/IDocumentHandler.cs:53-114`](file:///tmp/officecli-research/src/officecli/Core/IDocumentHandler.cs) | 3-layer interface |
| [`src/officecli/Core/PathIndex.cs:23-30`](file:///tmp/officecli-research/src/officecli/Core/PathIndex.cs) | 1-based ↔ 0-based boundary |
| [`src/officecli/Core/PathIdentity.cs:35-66`](file:///tmp/officecli-research/src/officecli/Core/PathIdentity.cs) | Canonical path + SHA256 cho named-pipe |
| [`src/officecli/Core/CliException.cs:1-27`](file:///tmp/officecli-research/src/officecli/Core/CliException.cs) | Structured error class |
| [`src/officecli/Core/EditDistance.cs:17-39`](file:///tmp/officecli-research/src/officecli/Core/EditDistance.cs) | Damerau-Levenshtein |
| [`src/officecli/Core/OutputFormatter.cs:103-200`](file:///tmp/officecli-research/src/officecli/Core/OutputFormatter.cs) | JSON envelope |
| [`src/officecli/Core/OutputFormatter.cs:82-101`](file:///tmp/officecli-research/src/officecli/Core/OutputFormatter.cs) | WarningContext thread-static |
| [`src/officecli/Core/TrackingPropertyDictionary.cs:1-160`](file:///tmp/officecli-research/src/officecli/Core/TrackingPropertyDictionary.cs) | Handler-as-truth pattern |
| [`src/officecli/Core/SkillInstaller.cs:11-25`](file:///tmp/officecli-research/src/officecli/Core/SkillInstaller.cs) | 13 agent targets table |
| [`src/officecli/Core/SkillInstaller.cs:67-78`](file:///tmp/officecli-research/src/officecli/Core/SkillInstaller.cs) | SkillTriggers (1-line) |
| [`src/officecli/Core/ResidentServer.cs:1-100`](file:///tmp/officecli-research/src/officecli/Core/ResidentServer.cs) | In-process document cache |
| [`src/officecli/McpServer.cs:208-260`](file:///tmp/officecli-research/src/officecli/McpServer.cs) | 1 thin shell tool pattern |
| [`src/officecli/McpServer.cs:541-590`](file:///tmp/officecli-research/src/officecli/McpServer.cs) | load_skill lazy reference |
| [`src/officecli/Help/SchemaHelpLoader.cs:1-200`](file:///tmp/officecli-research/src/officecli/Help/SchemaHelpLoader.cs) | Runtime schema loader |
| [`src/officecli/Handlers/Word/WordHandler.Navigation.cs:714+`](file:///tmp/officecli-research/src/officecli/Handlers/Word/WordHandler.Navigation.cs) | Word path parser |
| [`src/officecli/Handlers/DocumentHandlerFactory.cs`](file:///tmp/officecli-research/src/officecli/Handlers/DocumentHandlerFactory.cs) | Per-extension handler dispatch |
| [`src/officecli/CommandBuilder.cs:126-150`](file:///tmp/officecli-research/src/officecli/CommandBuilder.cs) | Probe-then-spawn singleton guard |
| [`src/officecli/CommandBuilder.cs:1720-1779`](file:///tmp/officecli-research/src/officecli/CommandBuilder.cs) | SuggestPropertyWithDistance |
| [`schemas/help/_schema.json`](file:///tmp/officecli-research/schemas/help/_schema.json) | Meta-schema cho per-element schemas |
| [`schemas/README.md`](file:///tmp/officecli-research/schemas/README.md) | CI contract cho schema |
| [`SKILL.md`](file:///tmp/officecli-research/SKILL.md) | Tài liệu routing cho AI agent (~25KB) |
| [`sdk/node/index.js:96-110`](file:///tmp/officecli-research/sdk/node/index.js) | Named-pipe path builder |
| [`sdk/node/index.js:208-220`](file:///tmp/officecli-research/sdk/node/index.js) | Empty reply = don't retry |
| [`sdk/node/index.d.ts`](file:///tmp/officecli-research/sdk/node/index.d.ts) | SDK API surface |

---

## 14. Next steps

Repo clone ở `/tmp/officecli-research/`. Tài liệu này tóm tắt đủ patterns để quyết định hướng tiếp theo mà không cần đào thêm.

**Các bước tiếp theo có thể làm ngay (1 commit mỗi cái):**
1. Áp dụng P0 × 3 (CliException + JSON envelope + Levenshtein) — 2-3 ngày
2. Viết unit tests cho EditDistance, CliException, OutputFormatter (copy logic từ OfficeCLI sang JS)
3. Refactor `server/tools.js` để dùng CliException cho tất cả 18 tool kinds
4. Refactor `server/mcp-server.js` để dùng CliException khi wrap kết quả

**Các bước cần thảo luận thêm:**
- Có muốn giữ 18 MCP tools hay chuyển sang "1 thin shell tool"?
- Có muốn áp dụng resident pattern cho Ori AI không? (Trade-off: phức tạp + 1 process, nhưng latency thấp + consistency cao)
- Có muốn viết `server/schemas/<element>.json` cho từng Word element (paragraph, run, table, comment, ...)?

---

*Tài liệu này tổng hợp từ research ngày 2026-09-01. Khi OfficeCLI release version mới, cần cập nhật lại line references vì code có thể dịch chuyển.*
