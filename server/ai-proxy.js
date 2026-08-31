/**
 * AI Proxy Adapter (Universal AI Gateway)
 * Tương thích 100% với cả chuẩn OpenAI-compatible và Anthropic-compatible
 * Hỗ trợ DeepSeek (V3/R1), OpenAI (GPT-4o), Claude (3.5/3.7), Gemini, Ollama, OpenRouter, OneAPI...
 */

export class AiProxy {
  constructor(defaultEndpoint, defaultApiKey, defaultModel) {
    this.defaultEndpoint = defaultEndpoint || process.env.AI_API_ENDPOINT || "https://api.deepseek.com/chat/completions";
    this.defaultApiKey = defaultApiKey || process.env.AI_API_KEY || "";
    this.defaultModel = defaultModel || process.env.AI_MODEL || "deepseek-chat";
  }

  /**
   * Xác định endpoint là dạng Anthropic hay OpenAI
   */
  isAnthropicEndpoint(endpoint = "") {
    const ep = endpoint.toLowerCase();
    return ep.includes("anthropic") || ep.includes("/messages") || ep.endsWith("/v1/messages");
  }

  /**
   * Chuẩn bị Headers cho request
   */
  buildHeaders(endpoint, apiKey) {
    const key = apiKey || this.defaultApiKey;
    const isAnthropic = this.isAnthropicEndpoint(endpoint);

    if (isAnthropic) {
      return {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      };
    }

    return {
      "content-type": "application/json",
      "authorization": `Bearer ${key}`,
    };
  }

  /**
   * Chuẩn bị Payload cho request
   */
  buildPayload(endpoint, { messages, systemPrompt, model, stream = true, maxTokens = 4096, temperature = 0.7 }) {
    const isAnthropic = this.isAnthropicEndpoint(endpoint);
    const selectedModel = model || this.defaultModel;

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
      return payload;
    }

    // OpenAI-compatible payload (DeepSeek, OpenAI, Gemini, Ollama, OpenRouter...)
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

    return {
      model: selectedModel,
      messages: fullMessages,
      max_tokens: maxTokens,
      stream: !!stream,
      temperature,
    };
  }

  /**
   * Xử lý request (tự động phân nhánh Stream SSE hoặc JSON phản hồi trực tiếp)
   */
  async handleRequest(reqBody, res) {
    const endpoint = reqBody.endpoint || this.defaultEndpoint;
    const apiKey = reqBody.apiKey || this.defaultApiKey;
    const stream = reqBody.stream !== false;
    const headers = this.buildHeaders(endpoint, apiKey);
    const payload = this.buildPayload(endpoint, reqBody);

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
      let errorBody = "";
      try {
        errorBody = await upstreamResponse.text();
      } catch (_) {}
      res.status(upstreamResponse.status).json({
        error: `Lỗi từ AI Endpoint (${upstreamResponse.status}): ${errorBody}`,
      });
      return;
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
              // Sự kiện từ OpenAI / DeepSeek
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
