/**
 * SSE Stream Client Module
 * Tiếp nhập luồng Server-Sent Events từ Backend, tách nhỏ token-by-token và bóc tách Thinking mode
 *
 * Endpoint có thể resolve qua 3 layer (stream-factory.js):
 *  1. StreamProxy (server-side, Bearer token)
 *  2. CORS Proxy (client-side rewrite)
 *  3. Direct (browser → provider với key localStorage)
 *  4. Default: server proxy /api/chat (zero-config fallback)
 */

import { docState } from "./doc-state.js?v=2.1.0";
import { resolveStreamEndpoint } from "./stream-factory.js?v=2.1.0";

export class SseStreamClient {
  constructor(endpointUrl = "/api/chat") {
    this.endpointUrl = endpointUrl;
    this.abortController = null;
  }

  /**
   * Lấy license headers từ localStorage (per-user)
   */
  getLicenseHeaders() {
    const lic = docState.loadLicense();
    if (!lic?.licenseKey || !lic?.deviceId) return {};
    return {
      "X-License-Key": lic.licenseKey,
      "X-Client-Id": lic.deviceId,
    };
  }

  /**
   * Bắt đầu gửi yêu cầu và đọc luồng stream
   *
   * SECURITY: KHÔNG gửi endpoint/apiKey/model từ client nữa.
   * Server lấy từ config. Chỉ gửi messages, systemPrompt, effort.
   */
  async streamChat({
    messages,
    systemPrompt,
    effort = "medium",
    onDelta,
    onThinking,
    onDone,
    onError,
  }) {
    // Hủy request trước đó nếu còn đang chạy
    this.abort();
    this.abortController = new AbortController();

    let fullText = "";
    let fullThinking = "";

    try {
      // Resolve endpoint: ưu tiên streamProxy/CORS/direct từ stream-factory
      let resolved = null;
      try {
        resolved = await resolveStreamEndpoint({ messages, systemPrompt, effort });
      } catch (_) {}
      const finalUrl = resolved?.url || this.endpointUrl;
      const extraHeaders = resolved?.headers || {};

      const response = await fetch(finalUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...extraHeaders,
          ...this.getLicenseHeaders(),
        },
        body: JSON.stringify({
          messages,
          systemPrompt,
          effort,
          stream: true,
        }),
        signal: this.abortController.signal,
      });

      if (!response.ok) {
        let errDetail = "";
        try {
          const errJson = await response.json();
          errDetail = errJson.error || errJson.message || response.statusText;
        } catch (_) {
          errDetail = await response.text();
        }
        throw new Error(errDetail || `Lỗi HTTP ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const events = buffer.split("\n\n");
        buffer = events.pop(); // giữ lại phần chưa hoàn chỉnh

        for (const eventBlock of events) {
          const lines = eventBlock.split("\n");
          let eventType = "message";
          let dataStr = "";

          for (const line of lines) {
            if (line.startsWith("event: ")) {
              eventType = line.slice(7).trim();
            } else if (line.startsWith("data: ")) {
              dataStr = line.slice(6);
            }
          }

          if (!dataStr) continue;

          try {
            const parsed = JSON.parse(dataStr);

            if (eventType === "thinking") {
              const chunk = parsed.text || "";
              fullThinking += chunk;
              if (onThinking) onThinking(chunk, fullThinking);
            } else if (eventType === "delta") {
              const chunk = parsed.text || "";
              fullText += chunk;
              if (onDelta) onDelta(chunk, fullText);
            } else if (eventType === "done") {
              if (onDone) onDone(fullText, fullThinking);
              return { fullText, fullThinking };
            } else if (eventType === "error") {
              throw new Error(parsed.error || "Lỗi luồng stream AI");
            }
          } catch (e) {
            if (eventType === "error") throw e;
          }
        }
      }

      if (onDone) onDone(fullText, fullThinking);
      return { fullText, fullThinking };
    } catch (err) {
      if (err.name === "AbortError") {
        if (onDone) onDone(fullText, fullThinking, true);
        return { fullText, fullThinking, aborted: true };
      }
      if (onError) onError(err);
      throw err;
    } finally {
      this.abortController = null;
    }
  }

  /**
   * Ngắt luồng sinh hiện tại
   */
  abort() {
    if (this.abortController) {
      this.abortController.abort();
      this.abortController = null;
    }
  }

  isStreaming() {
    return !!this.abortController;
  }
}

export const sseClient = new SseStreamClient();
