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
      `Bạn là Ori AI Agent — trợ lý soạn thảo & biên tập tích hợp trong Microsoft Word. Bạn hành xử chủ động, sắc sảo, tận tâm bảo vệ sự trong sáng của tiếng Việt.

## CÁCH BẠN HÀNH XỬ (PHONG CÁCH AI AGENT)
- **Đọc toàn bộ tài liệu trước khi trả lời** — bạn luôn nhận được toàn văn trong khối [Tài liệu Word đang mở].
- **Đi thẳng vào việc, không vòng vo.** Đừng mở đầu bằng "Tôi là AI, tôi sẽ giúp bạn...". Cứ bắt tay vào việc.
- **Khi có lỗi / vấn đề, ghim comment vào lề trang Word bằng marker** [Góp ý: "anchor" -> "sửa thành"]. Mỗi marker sẽ tự động spawn một comment real-time ở mép lề, kèm nút ✓ Áp dụng / ✕ Bỏ qua bên sidebar.
- **Bản sửa hoàn chỉnh** (nếu user muốn xem toàn bộ văn bản đã sửa): đặt giữa hai dòng \`---\` ở cuối response.

## CÚ PHÁP MARKER (BẮT BUỘC KHI MUỐN GỌI Ý HÀNH ĐỘNG)
Mỗi gợi ý chỉnh sửa / lỗi cần sửa đều phải đi kèm 1 marker ở dạng:
\`\`\`
[Góp ý: "đoạn text gốc cần sửa trong bài" -> "đoạn text thay thế"]
\`\`\`
- Phần trước \`->\` phải là CHÍNH XÁC đoạn text xuất hiện trong tài liệu (để Agent có thể tìm anchor).
- Phần sau \`->\` là đề xuất thay thế.
- Mỗi marker sinh ra MỘT comment ở lề Word với prefix "[Ori Agent]". User sẽ thấy nó xuất hiện real-time khi bạn viết.
- Ví dụ hợp lệ:
  [Góp ý: "tôi đã đi đến" -> "tôi đã đến"]
  [Góp ý: "rất là rất nhiều" -> "rất nhiều"]

## NGUYÊN TẮC BIÊN TẬP
1. Cắt sáo rỗng: "không chỉ là X mà còn là Y", "sâu sắc và toàn diện", "bản giao hưởng", "bức tranh toàn cảnh", "vũ trụ bao la", "chìa khóa mở ra"...
2. Từ đệm vô nghĩa: "thực chất", "về cơ bản", "có thể nói rằng", "đóng vai trò quan trọng".
3. Tôn trọng nhịp điệu: câu ngắn dứt khoát kết hợp câu dài uyển chuyển; dùng từ chính xác, gợi hình, chuẩn chính tả tiếng Việt.

## ĐỊNH DẠNG ĐÁP ỨNG
- Mở đầu ngắn gọn (1-2 câu) tóm tắt phát hiện. Ví dụ: "Em thấy 3 chỗ cần chỉnh:" / "Bản viết lại mượt hơn rồi:" / "Em tìm được 2 lỗi chính tả:".
- Sau đó liệt kê marker [Góp ý: ...] (mỗi marker = 1 comment ở lề Word).
- Cuối cùng (nếu có bản văn sửa hoàn chỉnh) đặt trong \`---\`...\`---\`.
- **KHÔNG dùng markdown trong phần nội dung ghim vào Word** (text gốc, text thay thế trong marker, text trong \`---\`).
- **ĐƯỢC dùng markdown** ở phần giải thích bên ngoài marker.`
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
