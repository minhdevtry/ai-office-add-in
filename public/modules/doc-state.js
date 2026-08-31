/**
 * Document State Persistence Module
 * Tự động lưu và phục hồi lịch sử chat, giọng văn cá nhân theo từng file Word (.docx)
 *
 * SECURITY: KHÔNG lưu API key / endpoint / model ở đây nữa (đã chuyển sang server-side).
 * Chỉ lưu:
 *  - HISTORY: lịch sử chat (per-file)
 *  - HOUSE_VOICE: mẫu giọng văn cá nhân (per-file)
 *  - SYS_PROMPT: custom system prompt (per-file)
 *  - EFFORT: thinking effort (per-user, chỉ localStorage)
 *  - LICENSE: thông tin license (per-user, chỉ localStorage)
 */

/* global Office */

const SETTINGS_KEYS = {
  HISTORY: "ai_word.history",
  HOUSE_VOICE: "ai_word.house_voice",
  SYS_PROMPT: "ai_word.system_prompt",
};

// Chỉ dùng localStorage — KHÔNG lưu vào document.settings
// (vì những thứ này là per-user-per-machine, không phải per-document)
const LOCAL_KEYS = {
  EFFORT: "ai_word.effort",
  LICENSE: "ai_word.license",
  DEVICE_ID: "ai_word.device_id",
};

export class DocumentStateManager {
  constructor() {
    this._saveTimer = null;
  }

  isSettingsAvailable() {
    return (
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.document &&
      Office.context.document.settings
    );
  }

  /**
   * Read value — ưu tiên document.settings, fallback localStorage.
   * Dùng cho dữ liệu per-document (history, house voice, system prompt).
   */
  get(key, defaultVal = null) {
    if (this.isSettingsAvailable()) {
      try {
        const val = Office.context.document.settings.get(key);
        return val !== null && val !== undefined ? val : defaultVal;
      } catch (_) {}
    }

    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : defaultVal;
    } catch (_) {
      return defaultVal;
    }
  }

  /**
   * Write value — ghi vào document.settings + localStorage fallback.
   * CHỈ dùng cho per-document data.
   */
  set(key, value) {
    if (this.isSettingsAvailable()) {
      try {
        Office.context.document.settings.set(key, value);
        Office.context.document.settings.saveAsync(() => {});
      } catch (_) {}
    }

    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (_) {}
  }

  /**
   * Read chỉ từ localStorage (per-user, KHÔNG lưu vào file).
   */
  getLocal(key, defaultVal = null) {
    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : defaultVal;
    } catch (_) {
      return defaultVal;
    }
  }

  /**
   * Write chỉ vào localStorage.
   */
  setLocal(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (_) {}
  }

  /**
   * Lưu trạng thái có debounce (tránh gọi I/O liên tục khi đang chat)
   */
  debouncedSave(key, value, delayMs = 500) {
    clearTimeout(this._saveTimer);
    this._saveTimer = setTimeout(() => {
      this.set(key, value);
    }, delayMs);
  }

  // ─── Per-document state (file .docx) ──────────────────────────────────

  loadHistory() {
    return this.get(SETTINGS_KEYS.HISTORY, []);
  }

  saveHistory(messages) {
    const truncated = messages.slice(-30);
    this.debouncedSave(SETTINGS_KEYS.HISTORY, truncated);
  }

  loadHouseVoice() {
    return this.get(SETTINGS_KEYS.HOUSE_VOICE, "");
  }

  saveHouseVoice(text) {
    this.set(SETTINGS_KEYS.HOUSE_VOICE, text);
  }

  loadSystemPrompt() {
    return this.get(
      SETTINGS_KEYS.SYS_PROMPT,
      "Bạn là một Trợ lý AI chuyên nghiệp hỗ trợ soạn thảo và biên tập văn bản tiếng Việt trong Microsoft Word."
    );
  }

  saveSystemPrompt(prompt) {
    this.set(SETTINGS_KEYS.SYS_PROMPT, prompt);
  }

  clearHistory() {
    this.set(SETTINGS_KEYS.HISTORY, []);
  }

  // ─── Per-user state (localStorage only) ───────────────────────────────

  /**
   * Effort level: "off" | "low" | "medium" | "high" | "xhigh"
   */
  loadEffort(defaultEffort = "medium") {
    return this.getLocal(LOCAL_KEYS.EFFORT, defaultEffort);
  }

  saveEffort(effort) {
    this.setLocal(LOCAL_KEYS.EFFORT, effort);
  }

  /**
   * License info: { email, licenseKey, deviceId, activatedAt }
   */
  loadLicense() {
    return this.getLocal(LOCAL_KEYS.LICENSE, null);
  }

  saveLicense(license) {
    this.setLocal(LOCAL_KEYS.LICENSE, license);
  }

  clearLicense() {
    try {
      localStorage.removeItem(LOCAL_KEYS.LICENSE);
    } catch (_) {}
  }

  /**
   * Device fingerprint (UUID v4) — tạo 1 lần, dùng cho license binding
   */
  loadDeviceId() {
    return this.getLocal(LOCAL_KEYS.DEVICE_ID, null);
  }

  saveDeviceId(id) {
    this.setLocal(LOCAL_KEYS.DEVICE_ID, id);
  }
}

export const docState = new DocumentStateManager();
export { LOCAL_KEYS };
