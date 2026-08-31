/**
 * WebSocket Bridge Client Module
 * Lắng nghe các lệnh điều khiển từ AI Agent ngoài (MCP Server) và thực thi trên Office.js
 */

import { wordBridge } from "./word-bridge.js?v=2.1.0";
import { docState } from "./doc-state.js?v=2.1.0";

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

  /**
   * Lấy license headers từ localStorage (per-user).
   * Trả về empty object nếu chưa có license (dev mode).
   */
  getLicenseHeaders() {
    const lic = docState.loadLicense();
    if (!lic?.licenseKey || !lic?.deviceId) return {};
    return {
      "X-License-Key": lic.licenseKey,
      "X-Client-Id": lic.deviceId,
    };
  }

  connect() {
    if (this.ws && (this.ws.readyState === 0 || this.ws.readyState === 1)) {
      return;
    }

    const url = this.getWsUrl();
    if (this.onStatusChange) this.onStatusChange("connecting");

    try {
      // Browser WebSocket API không hỗ trợ custom headers trong constructor.
      // Phải dùng Sec-WebSocket-Protocol trick hoặc attach query string.
      // Ở đây ta attach device-id vào query để server đọc được.
      const lic = docState.loadLicense();
      let fullUrl = url;
      if (lic?.licenseKey && lic?.deviceId) {
        const params = new URLSearchParams({
          k: lic.licenseKey,
          d: lic.deviceId,
        });
        fullUrl = `${url}?${params.toString()}`;
      }
      this.ws = new WebSocket(fullUrl);
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
        // Dùng native search + insertText/insertParagraph (gọn và chính xác hơn findReplace hack)
        return await wordBridge.insertAfterAnchor(op.anchor, op.text, {
          asParagraph: !!op.asParagraph,
          style: op.style || null,
        });
      }

      case "deleteText":
        return await wordBridge.deleteText(op.find, op.maxDeletions);

      case "setParagraphStyle":
        return await wordBridge.setParagraphStyle(op.anchor, op.style);

      case "insertComment":
        return await wordBridge.insertComment(op.commentText, op.anchor);

      case "getComments":
        return await wordBridge.getComments();

      case "acceptComment":
        return await wordBridge.acceptComment({
          anchorText: op.anchorText,
          replacementText: op.replacementText,
        });

      case "deleteComment":
        return await wordBridge.deleteComment({ anchorText: op.anchorText });

      case "replyComment":
        return await wordBridge.replyToComment({
          anchorText: op.anchorText,
          replyText: op.replyText,
        });

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
