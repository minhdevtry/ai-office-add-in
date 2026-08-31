/**
 * Main Express & WebSocket Server
 * Phục vụ đồng thời:
 * 1. Giao diện Taskpane Sidebar (/public)
 * 2. Endpoint Proxy SSE AI Stream (/api/chat)
 * 3. WebSocket Channel cho MCP Bridge (/ws)
 * 4. REST Endpoints (/tools, /op, /ops, /api/status, /api/license/*)
 */

import express from "express";
import http from "http";
import https from "https";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { WebSocketServer } from "ws";
import cors from "cors";
import { config, validateConfig, logConfigSummary } from "./config.js";
import { aiProxy } from "./ai-proxy.js";
import { clientManager, tools } from "./tools.js";
import { licenseRouter, requireLicense } from "./license/routes.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const publicDir = path.join(rootDir, "public");

const app = express();

// ─── CORS Whitelist ─────────────────────────────────────────────────────────
// Chỉ cho phép origins trong config.corsOrigins, hỗ trợ wildcard *.officeapps.live.com
const corsOriginPatterns = config.corsOrigins.map((o) => {
  if (o.includes("*")) {
    const escaped = o.replace(/[.+?^${}()|[\]\\]/g, "\\$&").replace(/\*/g, "[^.]+");
    return new RegExp(`^${escaped}$`);
  }
  return o;
});

app.use(
  cors({
    origin(origin, callback) {
      // Cho phép request không có Origin header (curl, server-to-server, Postman trong dev)
      if (!origin) return callback(null, true);
      for (const pat of corsOriginPatterns) {
        if (typeof pat === "string" && pat === origin) return callback(null, true);
        if (pat instanceof RegExp && pat.test(origin)) return callback(null, true);
      }
      return callback(new Error(`CORS blocked: ${origin}`));
    },
    credentials: false,
  })
);

app.use(express.json({ limit: "20mb" }));

// Serve static assets và giao diện Taskpane
app.use(express.static(publicDir));
app.use("/assets", express.static(path.join(rootDir, "assets")));

// ─── API Routes ─────────────────────────────────────────────────────────────

// License routes (public: /activate, /public-config; protected: /info, /devices, /heartbeat, /device/:id)
app.use("/api/license", licenseRouter);

// 1. Proxy SSE Chat Stream & AI API — YÊU CẦU LICENSE
app.post("/api/chat", requireLicense, async (req, res) => {
  try {
    // Gắn license info vào body để audit
    req.body._license = { email: req.license?.email, deviceId: req.deviceId };
    await aiProxy.handleRequest(req.body, res);
  } catch (err) {
    if (!res.headersSent) {
      res.status(500).json({ error: err.message });
    }
  }
});

// 2. Trạng thái hệ thống & kết nối
app.get("/api/status", (req, res) => {
  const clientStatus = clientManager.getStatus();
  res.json({
    status: "ok",
    version: "1.0.0",
    serverTime: new Date().toISOString(),
    aiEndpoint: config.ai.endpoint, // vẫn trả, nhưng đây là info public (đã set server-side)
    aiModel: config.ai.model,
    bridge: clientStatus,
  });
});

// 3. Danh mục công cụ
app.get("/tools", (req, res) => {
  res.json({
    name: "word-bridge",
    version: "1.0.0",
    tools,
  });
});

// 4. Thực thi 1 thao tác Word đơn lẻ qua REST — YÊU CẦU LICENSE
app.post("/op", requireLicense, async (req, res) => {
  const op = req.body;
  if (!op || !op.kind) {
    return res.status(400).json({ ok: false, error: "Thiếu trường 'kind' trong op" });
  }

  try {
    const result = await clientManager.dispatchOp(op, {
      targetClientId: req.body.clientId,
      timeoutMs: req.body.timeoutMs || 15000,
    });
    res.json({ ok: true, result });
  } catch (err) {
    res.status(502).json({
      ok: false,
      error: err.message,
      detail: err.detail || null,
    });
  }
});

// 5. Thực thi một chuỗi thao tác (Batch Ops) — YÊU CẦU LICENSE
app.post("/ops", requireLicense, async (req, res) => {
  const ops = Array.isArray(req.body) ? req.body : req.body.ops;
  const stopOnError = req.query.stopOnError === "true" || req.body.stopOnError === true;

  if (!Array.isArray(ops) || ops.length === 0) {
    return res.status(400).json({ ok: false, error: "Thiếu mảng các 'ops'" });
  }

  const results = [];
  for (const op of ops) {
    try {
      const result = await clientManager.dispatchOp(op);
      results.push({ ok: true, op: op.kind, result });
    } catch (err) {
      results.push({ ok: false, op: op.kind, error: err.message, detail: err.detail || null });
      if (stopOnError) break;
    }
  }

  const allOk = results.every((r) => r.ok);
  res.json({
    ok: allOk,
    total: ops.length,
    completed: results.length,
    results,
  });
});

// ─── Khởi tạo Server & WebSocket ────────────────────────────────────────────

let server;
if (config.useHttps) {
  const certDir = config.certDir;
  if (!certDir || !fs.existsSync(path.join(certDir, "localhost.crt"))) {
    console.error(
      `[FATAL] USE_HTTPS=true nhưng cert không tìm thấy ở ${certDir}. Chạy \`npm run certs\` trước.`
    );
    process.exit(1);
  }
  server = https.createServer(
    {
      key: fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
    },
    app
  );
} else {
  server = http.createServer(app);
}

const wss = new WebSocketServer({
  server,
  path: "/ws",
  // Verify client: nếu LICENSE_REQUIRED=true thì bắt buộc có license info
  // (qua query string ?k=...&d=... vì browser WebSocket API không gửi custom headers)
  verifyClient(info, callback) {
    if (!config.license.required) {
      return callback(true); // dev mode
    }
    const url = new URL(info.req.url, "http://localhost");
    const licenseKey = url.searchParams.get("k");
    const deviceId = url.searchParams.get("d");
    if (!licenseKey || !deviceId) {
      console.warn("[WS] Connection rejected: missing license in query string");
      return callback(false, 401, "Missing license");
    }
    callback(true);
  },
});

wss.on("connection", (ws, req) => {
  // Đọc license từ query string (browser WebSocket API không cho set custom headers)
  const url = new URL(req.url, "http://localhost");
  const licenseKey = url.searchParams.get("k") || req.headers["x-license-key"];
  const deviceId = url.searchParams.get("d") || req.headers["x-client-id"] || req.headers["x-device-id"];

  const clientId = clientManager.register(ws, {
    info: req.socket.remoteAddress,
    licenseKey,
    deviceId,
  });

  // Heartbeat ping/pong
  ws.isAlive = true;
  ws.on("pong", () => {
    ws.isAlive = true;
  });

  ws.on("message", (data) => {
    try {
      const msg = JSON.parse(data.toString());
      if (msg.type === "result") {
        clientManager.handleResult(msg);
      } else if (msg.type === "hello") {
        console.log(`[WebSocket] Word Taskpane đã kết nối (ID: ${clientId})`);
      }
    } catch (_) {}
  });
});

// Định kỳ ping kiểm tra kết nối WebSocket
const pingInterval = setInterval(() => {
  wss.clients.forEach((ws) => {
    if (ws.isAlive === false) return ws.terminate();
    ws.isAlive = false;
    ws.ping();
  });
}, 15000);

wss.on("close", () => {
  clearInterval(pingInterval);
});

// ─── Bootstrap ──────────────────────────────────────────────────────────────

validateConfig();
logConfigSummary();

server.listen(config.port, config.host, () => {
  console.log(`
══════════════════════════════════════════════════════════════
  🚀 AI WORD ADD-IN & MCP AGENT BRIDGE SERVER
  • Web / Taskpane URL : ${config.useHttps ? "https" : "http"}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}/index.html
  • WebSocket Bridge   : ${config.useHttps ? "wss" : "ws"}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}/ws
  • AI Chat Proxy      : ${config.useHttps ? "https" : "http"}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}/api/chat
  • Tool Catalog       : ${config.useHttps ? "https" : "http"}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}/tools
  • Status Check       : ${config.useHttps ? "https" : "http"}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}/api/status
══════════════════════════════════════════════════════════════
  `);
});
