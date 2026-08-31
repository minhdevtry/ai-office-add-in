/**
 * License Client Module
 * Quản lý lifecycle license phía client:
 *  - Lưu/đọc local license info (email, key, device-id)
 *  - Gọi server APIs: activate, info, devices, revoke, heartbeat
 *  - Tự động gắn X-License-Key + X-Client-Id headers vào fetch
 */

import { docState } from "./doc-state.js?v=2.1.0";
import { getOrCreateDeviceId } from "./device-fingerprint.js?v=2.1.0";

const API_BASE = "/api/license";

export class LicenseClient {
  constructor() {
    this._heartbeatTimer = null;
    this._onInvalidate = null; // callback khi license bị reject
  }

  // ─── Local license state ──────────────────────────────────────────────

  getLocalLicense() {
    return docState.loadLicense();
  }

  setLocalLicense(data) {
    docState.saveLicense(data);
  }

  clearLocalLicense() {
    docState.clearLicense();
  }

  isActivated() {
    const lic = this.getLocalLicense();
    return Boolean(lic?.licenseKey && lic?.deviceId && lic?.email);
  }

  getLicenseHeaders() {
    const lic = this.getLocalLicense();
    if (!lic?.licenseKey || !lic?.deviceId) return {};
    return {
      "X-License-Key": lic.licenseKey,
      "X-Client-Id": lic.deviceId,
    };
  }

  // ─── API calls ────────────────────────────────────────────────────────

  async fetchPublicConfig() {
    try {
      const res = await fetch(`${API_BASE}/public-config`);
      if (!res.ok) return { licenseRequired: false, aiDefaultEffort: "medium" };
      return await res.json();
    } catch (_) {
      return { licenseRequired: false, aiDefaultEffort: "medium" };
    }
  }

  async activate({ email, licenseKey, deviceName }) {
    const deviceId = getOrCreateDeviceId();
    try {
      const res = await fetch(`${API_BASE}/activate`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Client-Id": deviceId,
        },
        body: JSON.stringify({ email, licenseKey, deviceName }),
      });

      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) {
        const err = new Error(data.error || data.code || `HTTP ${res.status}`);
        err.code = data.code;
        err.detail = data.detail;
        err.status = res.status;
        throw err;
      }

      // Lưu local license
      this.setLocalLicense({
        email,
        licenseKey: licenseKey.toLowerCase(),
        deviceId,
        activatedAt: new Date().toISOString(),
      });

      return data;
    } catch (err) {
      // If server unreachable, allow dev/local mock activation
      if (err.message && err.message.includes("Failed to fetch")) {
        this.setLocalLicense({
          email,
          licenseKey: licenseKey.toLowerCase(),
          deviceId,
          activatedAt: new Date().toISOString(),
        });
        return { ok: true, isNewDevice: true, devMode: true };
      }
      throw err;
    }
  }

  async getInfo() {
    const res = await this._authFetch(`${API_BASE}/info`);
    return res.json();
  }

  async listDevices() {
    const res = await this._authFetch(`${API_BASE}/devices`);
    return res.json();
  }

  async revokeDevice(deviceId) {
    const res = await this._authFetch(`${API_BASE}/device/${deviceId}`, {
      method: "DELETE",
    });
    return res.json();
  }

  async heartbeat() {
    const res = await this._authFetch(`${API_BASE}/heartbeat`, {
      method: "POST",
    });
    return res.json();
  }

  async _authFetch(url, options = {}) {
    const lic = this.getLocalLicense();
    if (!lic?.licenseKey || !lic?.deviceId) {
      throw new Error("Chưa kích hoạt license");
    }
    const headers = {
      "Content-Type": "application/json",
      "X-License-Key": lic.licenseKey,
      "X-Client-Id": lic.deviceId,
      ...(options.headers || {}),
    };
    const res = await fetch(url, { ...options, headers });
    if (res.status === 401 || res.status === 403 || res.status === 404) {
      // License invalid → clear local + notify
      const data = await res.json().catch(() => ({}));
      this._handleInvalidate(data.code || "INVALID");
      const err = new Error(data.error || "License invalid");
      err.code = data.code;
      err.status = res.status;
      throw err;
    }
    return res;
  }

  // ─── Heartbeat loop ───────────────────────────────────────────────────

  /**
   * Bắt đầu heartbeat mỗi 5 phút.
   * Nếu license bị reject → gọi onInvalidate callback.
   */
  startHeartbeat(intervalMs = 5 * 60 * 1000) {
    this.stopHeartbeat();
    this._heartbeatTimer = setInterval(async () => {
      if (!this.isActivated()) return;
      try {
        await this.heartbeat();
      } catch (err) {
        console.warn("[License] heartbeat failed:", err.message);
      }
    }, intervalMs);
  }

  stopHeartbeat() {
    if (this._heartbeatTimer) {
      clearInterval(this._heartbeatTimer);
      this._heartbeatTimer = null;
    }
  }

  onInvalidate(callback) {
    this._onInvalidate = callback;
  }

  _handleInvalidate(code) {
    console.warn("[License] Invalidated:", code);
    this.clearLocalLicense();
    this.stopHeartbeat();
    if (this._onInvalidate) this._onInvalidate(code);
  }
}

export const licenseClient = new LicenseClient();
