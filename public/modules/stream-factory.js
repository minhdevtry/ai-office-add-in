/**
 * Composite Stream Factory (pi-for-word pattern)
 *
 * 3 layer fallback cho việc gọi AI streaming:
 *
 * 1. **StreamProxy (server-side)** — server giữ API key, browser gọi proxy URL + Bearer token
 * 2. **CORS Proxy (client-side rewrite)** — browser gửi qua CORS proxy URL
 * 3. **Direct** — browser gọi thẳng provider với API key lưu trong localStorage / IndexedDB
 *
 * Lợi ích:
 *  - User có thể dùng key riêng (OpenRouter, Anthropic direct, llama.cpp local) mà không cần server
 *  - Default (nếu user không cấu hình) = dùng server proxy hiện tại (zero-config)
 *
 * Tích hợp vào sse-client.js bằng cách thay thế URL endpoint bằng getEndpoint() của factory này.
 */


// Helper: nếu URL đã có path "/chat/completions" thì dùng nguyên, nếu đã có "/v1" thì nối "/chat/completions",
// nếu chỉ có host thì nối "/v1/chat/completions" (OpenAI-compatible convention).
function normalizeChatCompletionsUrl(rawUrl) {
  let u = rawUrl.replace(/\/$/, "");
  if (/\/chat\/completions(\?|$)/.test(u)) return u;
  if (/\/v1(\?|$)/.test(u)) return u + "/chat/completions";
  return u + "/v1/chat/completions";
}
const STORAGE_KEYS = {
  STREAM_PROXY: "ai_word.stream_proxy", // { enabled, url, token }
  CORS_PROXY: "ai_word.cors_proxy", // { enabled, url }
  PROVIDER_KEY: "ai_word.provider_key", // { provider, apiKey, baseURL }
};

const DEFAULT_PROVIDER = "openai";

/**
 * Lấy config stream proxy từ localStorage.
 * @returns {{ enabled: boolean, url: string, token: string }}
 */
function getStreamProxyConfig() {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return { enabled: false, url: "", token: "" };
    const raw = localStorage.getItem(STORAGE_KEYS.STREAM_PROXY);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return { enabled: false, url: "", token: "" };
}

/**
 * Lấy config CORS proxy từ localStorage.
 * @returns {{ enabled: boolean, url: string }}
 */
function getCorsProxyConfig() {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return { enabled: false, url: "" };
    const raw = localStorage.getItem(STORAGE_KEYS.CORS_PROXY);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return { enabled: false, url: "" };
}

/**
 * Lấy provider key từ localStorage.
 * @returns {{ provider: string, apiKey: string, baseURL: string }}
 */
function getProviderKeyConfig() {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return { provider: DEFAULT_PROVIDER, apiKey: "", baseURL: "" };
    const raw = localStorage.getItem(STORAGE_KEYS.PROVIDER_KEY);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return { provider: DEFAULT_PROVIDER, apiKey: "", baseURL: "" };
}

/**
 * Resolves endpoint URL + headers + body adjustments cho mỗi layer.
 * Trả về { url, headers, bodyMutator? } — call site dùng để rewrite fetch().
 *
 * @param {Object} req - request gốc { messages, systemPrompt, effort, stream, model }
 * @returns {Promise<{ url: string, headers: object, bodyMutator?: (body) => object }>}
 */
export async function resolveStreamEndpoint(req) {
  // Layer 1: StreamProxy (server-side, có Bearer token)
  const sp = getStreamProxyConfig();
  if (sp.enabled && sp.url && sp.token) {
    return {
      url: normalizeChatCompletionsUrl(sp.url),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${sp.token}`,
      },
    };
  }

  // Layer 2: CORS Proxy (client-side rewrite ?url=)
  const cp = getCorsProxyConfig();
  if (cp.enabled && cp.url) {
    const target = getProviderKeyConfig().baseURL || "https://api.openai.com/v1/chat/completions";
    // Xử lý URL có/không có query string:
    //   "https://proxy.com/"   → ".../?url=..."
    //   "https://proxy.com/?"  → ".../?url=..."  (bỏ "?" rỗng)
    //   "https://proxy.com/?x=1" → ".../?x=1&url=..."
    let base = cp.url;
    if (base.endsWith("?")) base = base.slice(0, -1);
    const sep = base.includes("?") ? "&" : "?";
    return {
      url: base + sep + "url=" + encodeURIComponent(target),
      headers: { "Content-Type": "application/json" },
    };
  }

  // Layer 3: Direct (browser → provider, dùng key từ localStorage)
  const pk = getProviderKeyConfig();
  if (pk.apiKey && pk.baseURL) {
    return {
      url: normalizeChatCompletionsUrl(pk.baseURL),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${pk.apiKey}`,
      },
    };
  }

  // Default: dùng server proxy hiện tại (zero-config fallback)
  return null;
}

/**
 * Save config từ UI Settings.
 */
export function saveStreamProxyConfig(cfg) {
  if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return;
  localStorage.setItem(STORAGE_KEYS.STREAM_PROXY, JSON.stringify(cfg));
}

export function saveCorsProxyConfig(cfg) {
  if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return;
  localStorage.setItem(STORAGE_KEYS.CORS_PROXY, JSON.stringify(cfg));
}

export function saveProviderKeyConfig(cfg) {
  if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return;
  localStorage.setItem(STORAGE_KEYS.PROVIDER_KEY, JSON.stringify(cfg));
}

export function getAllStreamConfigs() {
  return {
    streamProxy: getStreamProxyConfig(),
    corsProxy: getCorsProxyConfig(),
    providerKey: getProviderKeyConfig(),
  };
}
