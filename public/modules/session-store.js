/**
 * Session Store — IndexedDB-backed multi-turn session persistence
 *
 * Pattern từ word-GPT-Plus (api/checkpoints.ts) + pi-for-word (settings-storage.js).
 *
 * Mỗi session = 1 thread với thread_id, có nhiều checkpoint.
 * Cho phép:
 *  - Lưu message history + system prompt + config sau mỗi turn
 *  - Resume từ checkpoint gần nhất khi mở lại file Word
 *  - List sessions theo file Word (mỗi file có session riêng)
 *
 * API:
 *  - saveSession({ threadId, docPath, messages, systemPrompt, config })
 *  - loadLatestSession(docPath) → trả về session gần nhất
 *  - listSessions() → liệt kê tất cả sessions
 *  - deleteSession(threadId)
 *
 * Fallback to localStorage nếu IndexedDB không khả dụng (môi trường dev, iframe bị chặn).
 */

const DB_NAME = "ori_ai";
const DB_VERSION = 1;
const STORE_SESSIONS = "sessions";

let _dbPromise = null;

function getDB() {
  if (_dbPromise) return _dbPromise;
  _dbPromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB không khả dụng"));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_SESSIONS)) {
        const store = db.createObjectStore(STORE_SESSIONS, { keyPath: "threadId" });
        store.createIndex("docPath", "docPath", { unique: false });
        store.createIndex("updatedAt", "updatedAt", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return _dbPromise;
}

async function withStore(mode, fn) {
  const db = await getDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_SESSIONS, mode);
    const store = tx.objectStore(STORE_SESSIONS);
    const result = fn(store);
    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(tx.error);
  });
}

/**
 * Lưu / cập nhật 1 session.
 */
export async function saveSession({ threadId, docPath, messages, systemPrompt, config, title }) {
  const session = {
    threadId,
    docPath: docPath || "unknown",
    title: title || generateTitle(messages),
    messages: messages.slice(-50), // cap 50 messages gần nhất
    systemPrompt: systemPrompt || "",
    config: config || {},
    updatedAt: Date.now(),
  };

  try {
    await withStore("readwrite", (store) => store.put(session));
  } catch (err) {
    console.warn("[Ori AI] IndexedDB save failed, fallback to localStorage:", err.message);
    saveSessionToLocalStorage(session);
  }
  return session;
}

/**
 * Lấy session mới nhất cho 1 file Word cụ thể.
 */
export async function loadLatestSession(docPath) {
  try {
    const db = await getDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_SESSIONS, "readonly");
      const store = tx.objectStore(STORE_SESSIONS);
      const index = store.index("docPath");
      const req = index.getAll(docPath || "unknown");
      req.onsuccess = () => {
        const sessions = req.result || [];
        sessions.sort((a, b) => b.updatedAt - a.updatedAt);
        resolve(sessions[0] || null);
      };
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn("[Ori AI] IndexedDB load failed, fallback to localStorage:", err.message);
    return loadLatestSessionFromLocalStorage(docPath);
  }
}

/**
 * Liệt kê tất cả sessions, sắp xếp theo updatedAt giảm dần.
 */
export async function listSessions(limit = 50) {
  try {
    const db = await getDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_SESSIONS, "readonly");
      const store = tx.objectStore(STORE_SESSIONS);
      const req = store.getAll();
      req.onsuccess = () => {
        const all = req.result || [];
        all.sort((a, b) => b.updatedAt - a.updatedAt);
        resolve(all.slice(0, limit));
      };
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    return [];
  }
}

export async function deleteSession(threadId) {
  try {
    await withStore("readwrite", (store) => store.delete(threadId));
  } catch (_) {}
}

// ─── localStorage fallback (Office iframe có thể chặn localStorage) ───

const LS_PREFIX = "ai_word.session.";

// Memory fallback nếu localStorage bị chặn (Tracking Prevention trong Word iframe)
const _memoryStore = new Map();

function safeLocalStorageGet(key) {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) return _memoryStore.get(key) || null;
    return localStorage.getItem(key);
  } catch (_) {
    return _memoryStore.get(key) || null;
  }
}

function safeLocalStorageSet(key, value) {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) {
      _memoryStore.set(key, value);
      return;
    }
    localStorage.setItem(key, value);
  } catch (_) {
    // Office iframe có thể throw khi setItem do Tracking Prevention
    _memoryStore.set(key, value);
  }
}

function safeLocalStorageKeys(prefix) {
  try {
    if (typeof localStorage === "undefined" || (typeof window !== "undefined" && window.__ORI_NO_LS__)) {
      return Array.from(_memoryStore.keys()).filter((k) => k.startsWith(prefix));
    }
    return Object.keys(localStorage).filter((k) => k.startsWith(prefix));
  } catch (_) {
    return Array.from(_memoryStore.keys()).filter((k) => k.startsWith(prefix));
  }
}

function saveSessionToLocalStorage(session) {
  safeLocalStorageSet(LS_PREFIX + session.threadId, JSON.stringify(session));
}
function loadLatestSessionFromLocalStorage(docPath) {
  const keys = safeLocalStorageKeys(LS_PREFIX);
  const all = keys
    .map((k) => {
      try {
        return JSON.parse(safeLocalStorageGet(k) || "{}");
      } catch (_) {
        return null;
      }
    })
    .filter((s) => s && s.docPath === docPath)
    .sort((a, b) => b.updatedAt - a.updatedAt);
  return all[0] || null;
}

// ─── Helpers ─────────────────────────────────────────────────────────────

function generateTitle(messages) {
  const firstUser = (messages || []).find((m) => m.role === "user");
  if (!firstUser) return "Chat mới";
  const text = (firstUser.content || "").trim();
  return text.length > 50 ? text.slice(0, 50) + "…" : text;
}

/**
 * Tạo threadId mới (UUID v4).
 */
export function newThreadId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}
