/**
 * Main Express & WebSocket Server
 * Phục vụ đồng thời:
 * 1. Giao diện Taskpane Sidebar (/public)
 * 2. Endpoint Proxy SSE AI Stream (/api/chat)
 * 3. WebSocket Channel cho MCP Bridge (/ws)
 * 4. REST Endpoints (/tools, /op, /ops, /api/status)
 */

import express from "express";
import http from "http";
import https from "https";
import path from "path";
import { fileURLToPath } from "url";
import { WebSocketServer } from "ws";
import cors from "cors";
import dotenv from "dotenv";
import { aiProxy } from "./ai-proxy.js";
import { clientManager, tools } from "./tools.js";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const publicDir = path.join(rootDir, "public");

const app = express();
const port = parseInt(process.env.PORT || "3001", 10);
const host = process.env.HOST || "0.0.0.0";

app.use(cors());
app.use(express.json({ limit: "20mb" }));

// Serve static assets và giao diện Taskpane
app.use(express.static(publicDir));
app.use("/assets", express.static(path.join(rootDir, "assets")));

// ─── API Routes ─────────────────────────────────────────────────────────────

// 1. Proxy SSE Chat Stream & AI API
app.post("/api/chat", async (req, res) => {
  try {
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
    aiEndpoint: process.env.AI_API_ENDPOINT || "https://api.deepseek.com/chat/completions",
    aiModel: process.env.AI_MODEL || "deepseek-chat",
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

// 4. Thực thi 1 thao tác Word đơn lẻ qua REST
app.post("/op", async (req, res) => {
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

// 5. Thực thi một chuỗi thao tác (Batch Ops)
app.post("/ops", async (req, res) => {
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

const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: "/ws" });

wss.on("connection", (ws, req) => {
  const clientId = clientManager.register(ws, {
    info: req.socket.remoteAddress,
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

server.listen(port, host, () => {
  console.log(`
══════════════════════════════════════════════════════════════
  🚀 AI WORD ADD-IN & MCP AGENT BRIDGE SERVER
  • Web / Taskpane URL : http://localhost:${port}/index.html
  • WebSocket Bridge   : ws://localhost:${port}/ws
  • AI Chat Proxy      : http://localhost:${port}/api/chat
  • Tool Catalog       : http://localhost:${port}/tools
  • Status Check       : http://localhost:${port}/api/status
══════════════════════════════════════════════════════════════
  `);
});
