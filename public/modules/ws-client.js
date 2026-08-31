/**
 * WebSocket Bridge Client Module
 * Lắng nghe các lệnh điều khiển từ AI Agent ngoài (MCP Server) và thực thi trên Office.js
 */

import { wordBridge } from "./word-bridge.js";

export class WordWsClient {
  constructor() {
    this.ws = null;
    this.reconnectTimer = null;
    this.onStatusChange = null;
    this.onAgentActive = null;
  }

  getWsUrl() {
    const proto = window.location.protocol === "https:" ? "wss:" : "ws:";
    return `${proto}//${window.location.host}/ws`;
  }

  connect() {
    if (this.ws && (this.ws.readyState === 0 || this.ws.readyState === 1)) {
      return;
    }

    const url = this.getWsUrl();
    if (this.onStatusChange) this.onStatusChange("connecting");

    try {
      this.ws = new WebSocket(url);
    } catch (_) {
      this.scheduleReconnect();
      return;
    }

    this.ws.onopen = () => {
      if (this.onStatusChange) this.onStatusChange("connected");
      this.ws.send(
        JSON.stringify({
          type: "hello",
          kind: "word",
          info: "Microsoft Word Taskpane",
        })
      );
    };

    this.ws.onclose = () => {
      if (this.onStatusChange) this.onStatusChange("disconnected");
      this.scheduleReconnect();
    };

    this.ws.onerror = () => {
      if (this.onStatusChange) this.onStatusChange("error");
    };

    this.ws.onmessage = async (event) => {
      let msg;
      try {
        msg = JSON.parse(event.data);
      } catch (_) {
        return;
      }

      if (msg.type === "op") {
        const { id, op } = msg;
        if (this.onAgentActive) this.onAgentActive(true, op.kind);

        try {
          const result = await this.applyOp(op);
          this.ws.send(JSON.stringify({ type: "result", id, ok: true, result }));
        } catch (err) {
          this.ws.send(
            JSON.stringify({
              type: "result",
              id,
              ok: false,
              error: err.message || String(err),
            })
          );
        } finally {
          if (this.onAgentActive) this.onAgentActive(false);
        }
      }
    };
  }

  scheduleReconnect() {
    if (this.reconnectTimer) return;
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.connect();
    }, 2500);
  }

  /**
   * Bộ thực thi các lệnh từ Agent ngoài trên Office.js
   */
  async applyOp(op) {
    switch (op.kind) {
      case "ping":
        return { pong: true, time: new Date().toISOString() };

      case "getText":
        return {
          text: await wordBridge.getFullDocumentText(op.limit || 8000),
        };

      case "getParagraphs":
        return await wordBridge.getParagraphs(op.styleFilter, op.limit || 0, op.textLimit || 300);

      case "getSelection":
        return {
          selection: await wordBridge.getSelectedText(),
        };

      case "selectRange":
        return await wordBridge.selectRange(op.anchor, op.paragraphIndex);

      case "findReplace":
        return await wordBridge.findReplace(
          op.find,
          op.replace,
          op.matchCase,
          op.matchWholeWord,
          op.maxReplacements
        );

      case "insertAfterText": {
        if (!op.anchor || !op.text) throw new Error("Thiếu tham số 'anchor' hoặc 'text'");
        if (op.asParagraph) {
          await wordBridge.selectRange(op.anchor);
          await wordBridge.insertParagraphAfter(op.text);
          if (op.style) await wordBridge.setParagraphStyle(op.text, op.style);
        } else {
          await wordBridge.findReplace(op.anchor, `${op.anchor} ${op.text}`, true, false, 1);
        }
        return { ok: true };
      }

      case "deleteText":
        return await wordBridge.deleteText(op.find, op.maxDeletions);

      case "setParagraphStyle":
        return await wordBridge.setParagraphStyle(op.anchor, op.style);

      case "insertComment":
        return await wordBridge.insertComment(op.commentText, op.anchor);

      case "insertOoxml":
        return await wordBridge.insertOoxml(op.anchor, op.ooxml, op.location);

      case "getTrackedChanges":
        return await wordBridge.getTrackedChanges(op.timeoutMs);

      case "setTrackChanges":
        return await wordBridge.setTrackChanges(op.on);

      case "getMetadata":
        return await wordBridge.getDocumentMetadata();

      default:
        throw new Error(`Công cụ không được hỗ trợ: ${op.kind}`);
    }
  }
}

export const wsClient = new WordWsClient();
