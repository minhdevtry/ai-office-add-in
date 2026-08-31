/**
 * Device Fingerprint (Client)
 * Tạo UUID v4 ổn định cho mỗi máy, lưu trong localStorage.
 * Được dùng làm `X-Client-Id` cho license binding.
 */

import { docState } from "./doc-state.js?v=2.1.0";

function uuidv4() {
  // RFC 4122 v4
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  // Fallback
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

export function getOrCreateDeviceId() {
  let id = docState.loadDeviceId();
  if (!id) {
    id = uuidv4().toLowerCase();
    docState.saveDeviceId(id);
  }
  return id;
}

export function clearDeviceId() {
  docState.saveDeviceId(null);
}
