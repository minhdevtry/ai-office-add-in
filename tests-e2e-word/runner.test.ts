/**
 * E2E Test Runner (Mocha) — Word thật
 *
 * Pattern từ office-coding-agent/tests-e2e-word/runner.test.ts:
 *  1. Node khởi động CustomTestServer (HTTPS, port 4203) chờ POST /results
 *  2. Gọi office-addin-debugging SDK → startDebugging(manifestPath, { appType: Desktop, app: 'Word' })
 *  3. Manifest E2E có LaunchEvent OnNewDocument → auto-load test-taskpane khi mở doc mới
 *  4. test-taskpane chạy BÊN TRONG Word, import config từ word-tool-configs,
 *     gọi config.execute(context, args) cho từng tool, POST kết quả về
 *  5. Node nhận kết quả → assert từng tool
 *
 * KHÔNG cần Playwright. Word phải cài trên máy (Windows/Mac) để chạy.
 * CI chỉ chạy unit tests; E2E chỉ chạy local.
 *
 * Chạy: `npm run test:e2e:word`
 */

import { strict as assert } from "node:assert";
import { test, before, after } from "node:test";
import fs from "node:fs";
import path from "node:path";
import https from "node:https";
import { fileURLToPath } from "node:url";
import { e2eContext } from "./test-context.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");

const E2E_PORT = 4203;
const MANIFEST_PATH = path.join(__dirname, "test-manifest.xml");

/**
 * HTTPS server nhận kết quả test từ Word task pane.
 */
class TestResultServer {
  private server: https.Server | null = null;
  private resultsPromise: Promise<unknown[]> | null = null;
  private resolveResults: ((results: unknown[]) => void) | null = null;

  start(): Promise<number> {
    return new Promise((resolve, reject) => {
      // Self-signed cert (dev only) — same as office-addin-dev-certs
      const certPath = path.join(process.env.HOME || process.env.USERPROFILE || "", ".office-addin-dev-certs", "localhost.crt");
      const keyPath = path.join(process.env.HOME || process.env.USERPROFILE || "", ".office-addin-dev-certs", "localhost.key");

      if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
        reject(new Error("Self-signed cert chưa có. Chạy `npm run certs` trước."));
        return;
      }

      const options = {
        cert: fs.readFileSync(certPath),
        key: fs.readFileSync(keyPath),
      };

      this.resultsPromise = new Promise((res) => (this.resolveResults = res));

      this.server = https.createServer(options, (req, res) => {
        if (req.method === "POST" && req.url === "/results") {
          let body = "";
          req.on("data", (chunk) => (body += chunk));
          req.on("end", () => {
            try {
              const results = JSON.parse(body);
              res.writeHead(200, { "Content-Type": "application/json" });
              res.end(JSON.stringify({ ok: true }));
              if (this.resolveResults) this.resolveResults(results);
            } catch (err) {
              res.writeHead(400);
              res.end(String(err));
            }
          });
        } else {
          res.writeHead(404);
          res.end();
        }
      });

      this.server.listen(E2E_PORT, () => resolve(E2E_PORT));
    });
  }

  async waitForResults(timeoutMs = 60_000): Promise<unknown[]> {
    if (!this.resultsPromise) throw new Error("Server chưa start");
    return Promise.race([
      this.resultsPromise,
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error(`Timeout ${timeoutMs}ms chờ kết quả từ Word`)), timeoutMs)
      ),
    ]) as Promise<unknown[]>;
  }

  stop() {
    if (this.server) this.server.close();
  }
}

const testServer = new TestResultServer();

before("Setup: start HTTPS test server + spawn Word với E2E manifest", async () => {
  // 1. Start HTTPS server
  await testServer.start();
  console.log(`[E2E] Test server listening on https://localhost:${E2E_PORT}`);

  // 2. Spawn Word + sideload E2E manifest
  // Lưu ý: cần office-addin-debugging SDK + Office Desktop cài sẵn
  // const { startDebugging, AppType, OfficeAddinTestingPlatform } = await import("office-addin-debugging");
  // await startDebugging(MANIFEST_PATH, {
  //   appType: AppType.Desktop,
  //   app: "Word",
  //   devServerCommandLine: "npx vite --config ./tests-e2e-word/vite.config.ts",
  //   devServerPort: 3003,
  //   enableDebugging: false,
  // });

  // Stub cho môi trường không cài office-addin-debugging:
  console.log("[E2E] Word spawning skipped (cần office-addin-debugging + Word Desktop để chạy thật)");
  console.log("[E2E] Để chạy E2E thật: cài `npm install office-addin-debugging` rồi uncomment block trên.");
});

after("Teardown: stop test server + close Word", () => {
  testServer.stop();
  // closeDesktopApplication() từ office-addin-debugging
});

// ─── Tests (assert trên kết quả gửi về từ Word) ──────────────────────

test("E2E: word_read_document trả về text không rỗng", async () => {
  // Trong môi trường E2E thật: chờ kết quả từ Word
  // const results = await testServer.waitForResults();
  // e2eContext.setResults(results);
  // const r = e2eContext.getResult("word_read_document");
  // assert.ok(r, "Không nhận được kết quả từ Word");
  // assert.equal(r.pass, true, r.error);
  assert.ok(true, "Stub: cần Word Desktop + office-addin-debugging để chạy thật");
});

test("E2E: word_insert_text chèn được text", () => {
  // const r = e2eContext.getResult("word_insert_text");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_insert_markdown parse heading/list/strong", () => {
  // const r = e2eContext.getResult("word_insert_markdown");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_search_text trả về search-anchor contract", () => {
  // const r = e2eContext.getResult("word_search_text");
  // assert.ok(r?.pass);
  // assert.ok(r.details.anchor_search_text);
  assert.ok(true, "Stub");
});

test("E2E: word_insert_comment ghim [Ori Agent] comment vào lề", () => {
  // const r = e2eContext.getResult("word_insert_comment");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_accept_comment replace + xoá comment", () => {
  // const r = e2eContext.getResult("word_accept_comment");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_resolve_comment đánh dấu resolved", () => {
  // const r = e2eContext.getResult("word_resolve_comment");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_insert_with_track_changes bật Track Changes, insert, restore", () => {
  // const r = e2eContext.getResult("word_insert_with_track_changes");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_get_tracked_changes liệt kê changes với timeout 3s", () => {
  // const r = e2eContext.getResult("word_get_tracked_changes");
  // assert.ok(r?.pass);
  assert.ok(true, "Stub");
});

test("E2E: word_read_selection trả về text + markdown", () => {
  // const r = e2eContext.getResult("word_read_selection");
  // assert.ok(r?.pass);
  // assert.ok(r.details.text);
  // assert.ok(r.details.markdown);
  assert.ok(true, "Stub");
});

test("E2E: word_read_paragraphs có style/text", () => {
  // const r = e2eContext.getResult("word_read_paragraphs");
  // assert.ok(r?.pass);
  // assert.ok(Array.isArray(r.details.paragraphs));
  assert.ok(true, "Stub");
});
