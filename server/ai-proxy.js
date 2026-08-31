/**
 * AI Proxy Adapter (Universal AI Gateway)
 * Tương thích 100% với cả chuẩn OpenAI-compatible và Anthropic-compatible
 *
 * SECURITY: AI endpoint, API key, model đều lấy từ server-side config.
 * Client KHÔNG được override các giá trị này (tránh leak key qua .docx file).
 */

import { config, normalizeEffort, VALID_EFFORTS } from "./config.js";

// ─── Effort Mapping (5 mức → provider-specific param) ──────────────────────

/**
 * Map effort level → Anthropic thinking.budget_tokens
 * Anthropic yêu cầu: max_tokens > budget_tokens
 */
const EFFORT_BUDGET_TOKENS = {
  low: 1024,
  medium: 4096,
  high: 16384,
  xhigh: 32768,
};

/**
 * Effort levels supported by OpenAI reasoning models
 * (chỉ các model o-series, gpt-5.x mới có param này)
 */
const VALID_OPENAI_EFFORTS = new Set(["low", "medium", "high", "xhigh"]);

export class AiProxy {
  constructor() {
    this.defaultEndpoint = config.ai.endpoint;
    this.defaultApiKey = config.ai.apiKey;
    this.defaultModel = config.ai.model;
    this.defaultEffort = config.ai.defaultEffort;
  }

  /**
   * Xác định endpoint là dạng Anthropic hay OpenAI-compatible
   */
  isAnthropicEndpoint(endpoint = "") {
    const ep = endpoint.toLowerCase();
    return ep.includes("anthropic") || ep.includes("/messages") || ep.endsWith("/v1/messages");
  }

  /**
   * Chuẩn bị Headers cho request
   */
  buildHeaders(endpoint, apiKey) {
    const isAnthropic = this.isAnthropicEndpoint(endpoint);

    if (isAnthropic) {
      return {
        "content-type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      };
    }

    return {
      "content-type": "application/json",
      authorization: `Bearer ${apiKey}`,
    };
  }

  /**
   * Chuẩn bị Payload cho request — thêm param `effort` để map sang
   * Anthropic `thinking.budget_tokens` hoặc OpenAI `reasoning_effort`.
   */
  buildPayload(endpoint, { messages, systemPrompt, model, stream = true, maxTokens = 4096, temperature = 0.7, effort = null }) {
    const isAnthropic = this.isAnthropicEndpoint(endpoint);
    const selectedModel = model || this.defaultModel;
    const safeEffort = VALID_EFFORTS.includes(effort) ? effort : this.defaultEffort;

    if (isAnthropic) {
      const payload = {
        model: selectedModel,
        messages: messages
          .filter((m) => m.role === "user" || m.role === "assistant")
          .map((m) => ({
            role: m.role,
            content: typeof m.content === "string" ? m.content : JSON.stringify(m.content),
          })),
        max_tokens: maxTokens,
        stream: !!stream,
        temperature,
      };
      if (systemPrompt) {
        payload.system = systemPrompt;
      }

      // Map effort → Anthropic thinking
      if (safeEffort !== "off") {
        const budgetTokens = EFFORT_BUDGET_TOKENS[safeEffort] || EFFORT_BUDGET_TOKENS.medium;
        payload.thinking = { type: "enabled", budget_tokens: budgetTokens };
        // Anthropic yêu cầu: max_tokens > budget_tokens
        if (payload.max_tokens <= budgetTokens) {
          payload.max_tokens = budgetTokens + 1024;
        }
      }
      return payload;
    }

    // OpenAI-compatible payload (DeepSeek, OpenAI, Gemini, Ollama, OpenRouter, MiniMax...)
    const fullMessages = [];
    if (systemPrompt) {
      fullMessages.push({ role: "system", content: systemPrompt });
    }
    for (const msg of messages) {
      if (msg.role === "user" || msg.role === "assistant") {
        fullMessages.push({
          role: msg.role,
          content: typeof msg.content === "string" ? msg.content : JSON.stringify(msg.content),
        });
      }
    }

    const payload = {
      model: selectedModel,
      messages: fullMessages,
      max_tokens: maxTokens,
      stream: !!stream,
      temperature,
    };

    // Map effort → OpenAI reasoning_effort
    // Chỉ gửi khi effort hợp lệ với OpenAI spec và khác "off"
    if (safeEffort !== "off" && VALID_OPENAI_EFFORTS.has(safeEffort)) {
      payload.reasoning_effort = safeEffort;
    }

    return payload;
  }

  /**
   * Xử lý request — lấy endpoint/apiKey/model từ SERVER CONFIG (KHÔNG từ client)
   */
  async handleRequest(reqBody, res) {
    // SECURITY: Luôn lấy từ server config, bỏ qua mọi giá trị client gửi lên
    const endpoint = this.defaultEndpoint;
    const apiKey = this.defaultApiKey;
    const model = this.defaultModel;

    const stream = reqBody.stream !== false;
    const effort = reqBody.effort || this.defaultEffort;
    const maxTokens = reqBody.maxTokens || config.ai.maxTokens;
    const temperature = reqBody.temperature ?? config.ai.temperature;

    const headers = this.buildHeaders(endpoint, apiKey);
    const payload = this.buildPayload(endpoint, {
      messages: reqBody.messages || [],
      systemPrompt: reqBody.systemPrompt,
      model,
      stream,
      maxTokens,
      temperature,
      effort,
    });

    let upstreamResponse;
    try {
      upstreamResponse = await fetch(endpoint, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
      });
    } catch (err) {
      res.status(502).json({
        error: `Không thể kết nối tới AI API Endpoint (${endpoint}): ${err.message}`,
      });
      return;
    }

    if (!upstreamResponse.ok) {
      // Nếu OpenAI-compatible provider reject vì `xhigh` không support
      // → retry 1 lần với effort thấp hơn
      if (
        effort === "xhigh" &&
        !this.isAnthropicEndpoint(endpoint) &&
        upstreamResponse.status >= 400 &&
        upstreamResponse.status < 500
      ) {
        const retryPayload = this.buildPayload(endpoint, {
          messages: reqBody.messages || [],
          systemPrompt: reqBody.systemPrompt,
          model,
          stream,
          maxTokens,
          temperature,
          effort: "high", // fallback từ xhigh → high
        });
        try {
          upstreamResponse = await fetch(endpoint, {
            method: "POST",
            headers,
            body: JSON.stringify(retryPayload),
          });
          if (!upstreamResponse.ok) {
            const errBody = await upstreamResponse.text().catch(() => "");
            res.status(upstreamResponse.status).json({
              error: `Lỗi từ AI Endpoint (${upstreamResponse.status}): ${errBody}`,
              note: "Đã retry với effort=high nhưng vẫn fail.",
            });
            return;
          }
          // Ghi log warning
          console.warn(
            `[AI-PROXY] ⚠️  Provider reject xhigh, đã retry với high. Endpoint: ${endpoint}`
          );
        } catch (e) {
          res.status(502).json({
            error: `Lỗi retry sau khi provider reject xhigh: ${e.message}`,
          });
          return;
        }
      } else {
        let errorBody = "";
        try {
          errorBody = await upstreamResponse.text();
        } catch (_) {}
        res.status(upstreamResponse.status).json({
          error: `Lỗi từ AI Endpoint (${upstreamResponse.status}): ${errorBody}`,
        });
        return;
      }
    }

    // Nếu không stream (ví dụ: test connection)
    if (!stream) {
      try {
        const json = await upstreamResponse.json();
        let replyText = "";
        if (this.isAnthropicEndpoint(endpoint)) {
          replyText = json.content?.filter((b) => b.type === "text").map((b) => b.text).join("") || "";
        } else {
          replyText = json.choices?.[0]?.message?.content || "";
        }
        res.json({ ok: true, text: replyText, raw: json });
        return;
      } catch (err) {
        res.status(500).json({ error: `Lỗi đọc phản hồi JSON: ${err.message}` });
        return;
      }
    }

    // Xử lý luồng SSE Stream
    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.setHeader("X-Accel-Buffering", "no");
    res.flushHeaders?.();

    const isAnthropic = this.isAnthropicEndpoint(endpoint);
    const reader = upstreamResponse.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop(); // giữ lại phần chưa đủ 1 dòng

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || trimmed.startsWith(":") || !trimmed.startsWith("data:")) continue;

          const dataStr = trimmed.slice(5).trim();
          if (dataStr === "[DONE]") {
            res.write(`event: done\ndata: {}\n\n`);
            continue;
          }

          try {
            const parsed = JSON.parse(dataStr);

            if (isAnthropic) {
              // Sự kiện từ Anthropic (Messages API)
              if (parsed.type === "content_block_delta" && parsed.delta) {
                if (parsed.delta.type === "text_delta" && parsed.delta.text) {
                  res.write(`event: delta\ndata: ${JSON.stringify({ text: parsed.delta.text })}\n\n`);
                } else if (parsed.delta.type === "thinking_delta" && parsed.delta.thinking) {
                  res.write(`event: thinking\ndata: ${JSON.stringify({ text: parsed.delta.thinking })}\n\n`);
                }
              }
            } else {
              // Sự kiện từ OpenAI / DeepSeek / MiniMax
              const choice = parsed.choices?.[0];
              if (choice?.delta) {
                // Hỗ trợ thinking từ DeepSeek R1 (reasoning_content)
                if (choice.delta.reasoning_content) {
                  res.write(`event: thinking\ndata: ${JSON.stringify({ text: choice.delta.reasoning_content })}\n\n`);
                }
                if (choice.delta.content) {
                  res.write(`event: delta\ndata: ${JSON.stringify({ text: choice.delta.content })}\n\n`);
                }
              }
            }
          } catch (_) {
            // bỏ qua dòng json lỗi
          }
        }
      }

      // Xử lý buffer còn lại
      if (buffer.trim().startsWith("data:")) {
        const dataStr = buffer.trim().slice(5).trim();
        if (dataStr !== "[DONE]") {
          try {
            const parsed = JSON.parse(dataStr);
            const content = parsed.choices?.[0]?.delta?.content || parsed.delta?.text;
            if (content) {
              res.write(`event: delta\ndata: ${JSON.stringify({ text: content })}\n\n`);
            }
          } catch (_) {}
        }
      }

      res.write(`event: done\ndata: {}\n\n`);
      res.end();
    } catch (err) {
      try {
        res.write(`event: error\ndata: ${JSON.stringify({ error: err.message })}\n\n`);
      } catch (_) {}
      res.end();
    } finally {
      reader.releaseLock();
    }
  }
}

export const aiProxy = new AiProxy();

// Re-export để test/export từ module khác
export { EFFORT_BUDGET_TOKENS, VALID_OPENAI_EFFORTS };
