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

// Safe memory fallback khi trình duyệt chặn localStorage (Tracking Prevention / Third-party cookies)
const memoryStore = new Map();
let isStorageWorking = null;

function checkStorage() {
  if (isStorageWorking !== null) return isStorageWorking;
  try {
    const testKey = "__ori_storage_test__";
    window.localStorage.setItem(testKey, "1");
    window.localStorage.removeItem(testKey);
    isStorageWorking = true;
  } catch (_) {
    isStorageWorking = false;
  }
  return isStorageWorking;
}

function safeStorageGet(key) {
  if (checkStorage()) {
    try {
      return window.localStorage.getItem(key);
    } catch (_) {}
  }
  return memoryStore.get(key) || null;
}

function safeStorageSet(key, value) {
  if (checkStorage()) {
    try {
      window.localStorage.setItem(key, value);
      return;
    } catch (_) {}
  }
  memoryStore.set(key, value);
}

function safeStorageRemove(key) {
  if (checkStorage()) {
    try {
      window.localStorage.removeItem(key);
      return;
    } catch (_) {}
  }
  memoryStore.delete(key);
}

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
      const stored = safeStorageGet(key);
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
      safeStorageSet(key, JSON.stringify(value));
    } catch (_) {}
  }

  /**
   * Read chỉ từ localStorage (per-user, KHÔNG lưu vào file).
   */
  getLocal(key, defaultVal = null) {
    try {
      const stored = safeStorageGet(key);
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
      safeStorageSet(key, JSON.stringify(value));
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
      `Bạn là Ori AI Agent — trợ lý soạn thảo & biên tập cao cấp tích hợp trực tiếp trong Microsoft Word theo mô hình AI Coding Agent (tương tự Cursor Composer / Claude Code dành cho văn bản).

## NGUYÊN TẮC HOẠT ĐỘNG
1. **Workspace Context**: Bạn luôn đọc và nắm trọn toàn bộ nội dung tài liệu Word đang mở trong khối [Tài liệu Word đang mở].
2. **Phong cách AI Agent**: Sắc sảo, súc tích, đi thẳng vào giải pháp. Không chào hỏi hay rào đón rườm rà.
3. **Cơ chế Đề xuất Sửa đổi (Patch / Margin Comments)**:
   - Khi phát hiện lỗi chính tả, sai dấu tiếng Việt, câu từ lủng củng, sáo rỗng, hoặc cần viết lại một đoạn: BẮT BUỘC dùng marker cú pháp sau cho từng điểm sửa:
     [Góp ý: "đoạn text gốc chính xác trong bài" -> "đoạn text thay thế hoàn chỉnh"]
   - Hệ thống sẽ tự động ghim Margin Comment [Ori Agent] lên mép lề Word và tạo thẻ Diff tương tác trong Chat để người dùng duyệt (Accept/Reject).
   - "đoạn text gốc" PHẢI xuất hiện nguyên văn chính xác trong tài liệu để hệ thống tìm được vị trí (anchor).
   - "đoạn text thay thế" là câu văn hoàn chỉnh sau khi đã sửa.
4. **Nội dung mới hoặc Văn bản tổng thể**: Nếu người dùng yêu cầu viết một đoạn mới hoặc bản sửa hoàn chỉnh toàn bài, hãy đặt đoạn văn bản đó giữa hai dòng phân cách \`---\` ở cuối bài để người dùng chèn 1-chạm vào Word.`
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
      safeStorageRemove(LOCAL_KEYS.LICENSE);
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
export { LOCAL_KEYS, safeStorageGet, safeStorageSet, safeStorageRemove };
