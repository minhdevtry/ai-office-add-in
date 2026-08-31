/**
 * Document State Persistence Module
 * Tự động lưu và phục hồi lịch sử chat, giọng văn cá nhân và cài đặt theo từng file Word (.docx)
 */

/* global Office */

const SETTINGS_KEYS = {
  HISTORY: "ai_word.history",
  HOUSE_VOICE: "ai_word.house_voice",
  SYS_PROMPT: "ai_word.system_prompt",
  CONFIG: "ai_word.config",
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

  get(key, defaultVal = null) {
    if (this.isSettingsAvailable()) {
      try {
        const val = Office.context.document.settings.get(key);
        return val !== null && val !== undefined ? val : defaultVal;
      } catch (_) {}
    }

    // LocalStorage fallback
    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : defaultVal;
    } catch (_) {
      return defaultVal;
    }
  }

  set(key, value) {
    if (this.isSettingsAvailable()) {
      try {
        Office.context.document.settings.set(key, value);
        Office.context.document.settings.saveAsync(() => {});
      } catch (_) {}
    }

    // LocalStorage fallback
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

  // --- Convenience Methods ---

  loadHistory() {
    return this.get(SETTINGS_KEYS.HISTORY, []);
  }

  saveHistory(messages) {
    // Giới hạn lưu 30 tin nhắn gần nhất để không làm nặng file
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

  loadConfig() {
    return this.get(SETTINGS_KEYS.CONFIG, {
      endpoint: "",
      apiKey: "",
      model: "",
      enableThinking: true,
      autoTrackChanges: false,
    });
  }

  saveConfig(cfg) {
    this.set(SETTINGS_KEYS.CONFIG, cfg);
  }

  clearHistory() {
    this.set(SETTINGS_KEYS.HISTORY, []);
  }
}

export const docState = new DocumentStateManager();
