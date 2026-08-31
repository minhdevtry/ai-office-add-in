/**
 * E2E Test Taskpane — chạy BÊN TRONG Word thật
 *
 * Pattern từ office-coding-agent/tests-e2e-word/src/test-taskpane.ts:
 *  1. Office.actions.associate('onWordLaunch', ...) auto-fire khi mở doc mới
 *  2. Gọi trực tiếp config.execute(context, args) từ word-tool-configs
 *  3. POST kết quả về Node test runner qua HTTPS POST /results
 *
 * KHÔNG mock Office.js — chạy thật trong Word.
 */

import { WORD_TOOL_CONFIGS, executeWordTool } from "../../public/modules/word-tool-configs/index.js";

const RESULTS_URL = "https://localhost:4203/results";

interface TestResult {
  name: string;
  pass: boolean;
  durationMs: number;
  error?: string;
  details?: Record<string, unknown>;
}

async function postResults(results: TestResult[]) {
  try {
    await fetch(RESULTS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(results),
    });
  } catch (err) {
    console.error("[E2E] Failed to post results:", err);
  }
}

async function runOne(name: string, fn: () => Promise<any>): Promise<TestResult> {
  const start = Date.now();
  try {
    const details = await fn();
    return { name, pass: true, durationMs: Date.now() - start, details };
  } catch (err: any) {
    return { name, pass: false, durationMs: Date.now() - start, error: err.message || String(err) };
  }
}

async function setupDocument() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.clear();
    body.insertHtml(
      "<h1>E2E Test Document</h1>" +
        "<p>Đây là đoạn văn test đầu tiên có chữ <strong>in đậm</strong> và <em>in nghiêng</em>.</p>" +
        "<h2>Heading 2</h2>" +
        "<p>Đoạn văn thứ hai với danh sách:</p>" +
        "<ul><li>Mục 1</li><li>Mục 2</li><li>Mục 3</li></ul>" +
        "<p>Anchor text for search: xyz_unique_marker_123</p>",
      Word.InsertLocation.start
    );
    await context.sync();
  });
}

async function runAllTests(): Promise<TestResult[]> {
  const results: TestResult[] = [];

  await setupDocument();

  // 1. word_read_document
  results.push(
    await runOne("word_read_document", async () => {
      const r = await executeWordTool("word_read_document", { limit: 50000 });
      if (!r.text || r.text.length < 10) throw new Error("Text rỗng hoặc quá ngắn");
      return { text_length: r.text.length };
    })
  );

  // 2. word_read_paragraphs
  results.push(
    await runOne("word_read_paragraphs", async () => {
      const r = await executeWordTool("word_read_paragraphs", { limit: 10 });
      if (!Array.isArray(r.paragraphs)) throw new Error("paragraphs không phải array");
      return { count: r.total };
    })
  );

  // 3. word_read_selection
  results.push(
    await runOne("word_read_selection", async () => {
      // Select all first via Range
      await Word.run(async (context) => {
        const sel = context.document.body;
        sel.select();
        await context.sync();
      });
      const r = await executeWordTool("word_read_selection", { format: "markdown" });
      if (!r.markdown || r.markdown.length < 5) throw new Error("markdown rỗng");
      return { text_length: r.text.length, markdown_length: r.markdown.length };
    })
  );

  // 4. word_search_text (trả anchor contract)
  results.push(
    await runOne("word_search_text", async () => {
      const r = await executeWordTool("word_search_text", {
        searchText: "xyz_unique_marker_123",
      });
      if (r.matchCount < 1) throw new Error("Không tìm thấy anchor");
      if (!r.anchor_search_text) throw new Error("Thiếu anchor_search_text trong contract");
      return { matchCount: r.matchCount, anchor: r.anchor_search_text };
    })
  );

  // 5. word_insert_text
  results.push(
    await runOne("word_insert_text", async () => {
      return await executeWordTool("word_insert_text", {
        text: "\n\nInserted by E2E test\n",
        location: "end",
      });
    })
  );

  // 6. word_insert_markdown
  results.push(
    await runOne("word_insert_markdown", async () => {
      return await executeWordTool("word_insert_markdown", {
        markdown: "## Heading 2 mới\n\nĐoạn **in đậm** và *in nghiêng*.\n\n- Item 1\n- Item 2\n",
        where: "end",
      });
    })
  );

  // 7. word_insert_comment (ghim [Ori Agent] comment)
  results.push(
    await runOne("word_insert_comment", async () => {
      return await executeWordTool("word_insert_comment", {
        anchorText: "Anchor text for search",
        commentText: "Đây là gợi ý test từ Ori Agent",
      });
    })
  );

  // 8. word_accept_comment (replace + xoá)
  results.push(
    await runOne("word_accept_comment", async () => {
      return await executeWordTool("word_accept_comment", {
        anchorText: "Anchor text for search: xyz_unique_marker_123",
        replacementText: "Anchor đã được sửa bởi E2E test",
      });
    })
  );

  // 9. word_resolve_comment (Word 1.4 API)
  results.push(
    await runOne("word_resolve_comment", async () => {
      // Insert 1 comment mới trước
      await executeWordTool("word_insert_comment", {
        anchorText: "Đoạn văn thứ hai",
        commentText: "Comment cần resolve",
      });
      return await executeWordTool("word_resolve_comment", {
        anchorText: "Đoạn văn thứ hai",
      });
    })
  );

  // 10. word_insert_with_track_changes
  results.push(
    await runOne("word_insert_with_track_changes", async () => {
      return await executeWordTool("word_insert_with_track_changes", {
        text: "\nTracked changes insert từ E2E test.\n",
        location: "after",
      });
    })
  );

  // 11. word_get_tracked_changes (timeout 3s)
  results.push(
    await runOne("word_get_tracked_changes", async () => {
      return await executeWordTool("word_get_tracked_changes", { timeoutMs: 3000 });
    })
  );

  return results;
}

async function bootstrap() {
  if (typeof Office === "undefined" || !Word) {
    console.error("[E2E] Office.js không khả dụng — chỉ chạy được trong Word.");
    return;
  }

  // Bind LaunchEvent: chạy ngay khi Word mở tài liệu mới
  if (Office.actions) {
    Office.actions.associate("onWordLaunch", async (event: Office.AddinCommands.Event) => {
      try {
        const results = await runAllTests();
        await postResults(results);
        console.log(`[E2E] Hoàn thành: ${results.filter((r) => r.pass).length}/${results.length} pass`);
      } catch (err: any) {
        await postResults([
          { name: "fatal", pass: false, durationMs: 0, error: err.message || String(err) },
        ]);
      } finally {
        event.completed();
      }
    });
  }

  // Nếu LaunchEvent không fire (Word cũ), chạy luôn sau onReady
  await Office.onReady();
  try {
    const results = await runAllTests();
    await postResults(results);
  } catch (err: any) {
    console.error("[E2E] Fatal:", err);
  }
}

bootstrap();
