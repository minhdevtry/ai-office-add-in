/**
 * Centralized Configuration Loader
 * Toàn bộ config được load từ environment variables với fallback defaults.
 * Mọi thay đổi về endpoint, API key, model, effort... đều đi qua đây.
 */

import "dotenv/config";

// ─── Helper functions ───────────────────────────────────────────────────────

function envInt(key, defaultValue) {
  const v = process.env[key];
  if (v === undefined || v === "") return defaultValue;
  const n = parseInt(v, 10);
  return Number.isNaN(n) ? defaultValue : n;
}

function envBool(key, defaultValue) {
  const v = process.env[key];
  if (v === undefined || v === "") return defaultValue;
  return v.toLowerCase() === "true" || v === "1";
}

function envList(key, defaultValue) {
  const v = process.env[key];
  if (!v) return defaultValue;
  return v.split(",").map((s) => s.trim()).filter(Boolean);
}

function envString(key, defaultValue) {
  const v = process.env[key];
  return v === undefined || v === "" ? defaultValue : v;
}

// ─── Config schema ──────────────────────────────────────────────────────────

/**
 * Allowed effort levels (case-insensitive at read time).
 * Server default = "medium", user can override via env AI_DEFAULT_EFFORT.
 */
const VALID_EFFORTS = ["off", "low", "medium", "high", "xhigh"];

function normalizeEffort(v) {
  if (!v) return "medium";
  const lower = String(v).toLowerCase();
  return VALID_EFFORTS.includes(lower) ? lower : "medium";
}

export const config = {
  // Server
  port: envInt("PORT", 3011),
  host: envString("HOST", "0.0.0.0"),
  publicUrl: envString("PUBLIC_URL", "https://word-bridge.2tocom.space"),
  useHttps: envBool("USE_HTTPS", false), // dev: false; Cloudflare Tunnel terminate TLS ở edge
  certDir: envString("OFFICE_ADDIN_CERT_DIR", ""),

  // AI Provider (server-side, client KHÔNG được override)
  ai: {
    endpoint: envString("AI_API_ENDPOINT", "https://aiapi.2tocom.space/v1/chat/completions"),
    apiKey: envString("AI_API_KEY", "keykeykeykeykeykeykeykey"),
    model: envString("AI_MODEL", "MiniMax-M3"),
    defaultEffort: normalizeEffort(envString("AI_DEFAULT_EFFORT", "medium")),
    // Optional override
    maxTokens: envInt("AI_MAX_TOKENS", 4096),
    temperature: parseFloat(envString("AI_TEMPERATURE", "0.7")) || 0.7,
  },

  // Supabase (License DB)
  supabase: {
    url: envString("SUPABASE_URL", ""),
    serviceKey: envString("SUPABASE_SERVICE_KEY", ""),
  },

  // CORS whitelist
  corsOrigins: envList("CORS_ORIGINS", [
    "https://localhost:3001",
    "http://localhost:3001",
    "https://*.officeapps.live.com", // Word Web domain wildcard (express will check via regex)
  ]),

  // License logic
  license: {
    required: envBool("LICENSE_REQUIRED", true),
    heartbeatTtlSec: envInt("LICENSE_HEARTBEAT_TTL", 3600),
    activateRateLimitPerMin: envInt("LICENSE_ACTIVATE_RATE_PER_MIN", 5),
  },
};

// ─── Validation & startup warnings ──────────────────────────────────────────

export function validateConfig() {
  const warnings = [];

  if (config.license.required) {
    if (!config.supabase.url || !config.supabase.serviceKey) {
      warnings.push(
        "[CONFIG] ⚠️  LICENSE_REQUIRED=true nhưng SUPABASE_URL / SUPABASE_SERVICE_KEY chưa set. License middleware sẽ reject mọi request."
      );
    }
  }

  if (!config.ai.apiKey) {
    warnings.push("[CONFIG] ⚠️  AI_API_KEY chưa set — chat sẽ fail khi gọi upstream.");
  }

  if (config.useHttps && !config.certDir) {
    warnings.push(
      "[CONFIG] ⚠️  USE_HTTPS=true nhưng OFFICE_ADDIN_CERT_DIR chưa set. Cần install cert trước khi start."
    );
  }

  for (const w of warnings) console.error(w);

  return warnings.length === 0;
}

// In ra summary lúc boot (che giấu API key)
export function logConfigSummary() {
  console.log(`
╔══════════════════════════════════════════════════════════════════╗
║  AI WORD ADD-IN & MCP AGENT BRIDGE — BOOT                       ║
╠══════════════════════════════════════════════════════════════════╣
║  Port           : ${config.port}                                     ║
║  HTTPS          : ${config.useHttps}                                  ║
║  AI Endpoint    : ${config.ai.endpoint}                  ║
║  AI Model       : ${config.ai.model}                                ║
║  AI Default     : effort=${config.ai.defaultEffort}, max=${config.ai.maxTokens} tokens, T=${config.ai.temperature}                ║
║  Supabase       : ${config.supabase.url ? config.supabase.url : "(chưa cấu hình)"}             ║
║  License Req.   : ${config.license.required}                                ║
║  CORS Origins   : ${config.corsOrigins.length} whitelisted                          ║
╚══════════════════════════════════════════════════════════════════╝`);
}

export { VALID_EFFORTS, normalizeEffort };
