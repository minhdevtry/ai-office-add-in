/**
 * License API Routes
 * Mounted tại /api/license/*
 *
 * Public routes (không cần license):
 *   POST   /activate
 *   GET    /public-config
 *
 * Protected routes (cần X-License-Key + X-Client-Id):
 *   GET    /info
 *   GET    /devices
 *   DELETE /device/:deviceId
 *   POST   /heartbeat
 */

import express from "express";
import rateLimit from "express-rate-limit";
import { licenseService } from "./license-service.js";
import { config } from "../config.js";
import { getClientIp, getUserAgent } from "./device-fingerprint.js";

export const licenseRouter = express.Router();

// ─── Rate limit cho /activate (chống brute force) ───────────────────────
const activateLimiter = rateLimit({
  windowMs: 60 * 1000, // 1 phút
  max: config.license.activateRateLimitPerMin,
  message: { ok: false, code: "RATE_LIMITED", error: "Quá nhiều lần thử. Vui lòng đợi 1 phút." },
  standardHeaders: true,
  legacyHeaders: false,
});

// Helper: lấy license key + client id từ header
function extractLicenseHeaders(req) {
  return {
    licenseKey: req.headers["x-license-key"],
    deviceId: req.headers["x-client-id"] || req.headers["x-device-id"],
  };
}

// ─── PUBLIC: Public config (cho client lấy model + default effort) ─────
licenseRouter.get("/public-config", (req, res) => {
  res.json({
    ok: true,
    aiModel: config.ai.model,
    aiDefaultEffort: config.ai.defaultEffort,
    licenseRequired: config.license.required,
    // KHÔNG trả endpoint / apiKey / license info cá nhân
  });
});

// ─── PUBLIC: Activate (enroll device) ────────────────────────────────────
licenseRouter.post("/activate", activateLimiter, async (req, res) => {
  const { email, licenseKey, deviceName } = req.body || {};
  const clientId = req.headers["x-client-id"] || req.headers["x-device-id"];
  const ip = getClientIp(req);
  const ua = getUserAgent(req);

  try {
    const result = await licenseService.activate({
      email,
      licenseKey,
      clientId,
      deviceName: deviceName || parseDeviceNameFromUA(ua),
      ip,
      ua,
    });

    if (!result.ok) {
      const status = result.httpStatus || 400;
      return res.status(status).json({
        ok: false,
        code: result.code,
        error: mapCodeToVN(result.code, result.detail),
        detail: result.detail || null,
      });
    }

    res.json({
      ok: true,
      isNewDevice: result.isNewDevice,
      license: result.license,
      device: result.device,
    });
  } catch (err) {
    console.error("[License] /activate error:", err);
    res.status(500).json({ ok: false, code: "INTERNAL_ERROR", error: err.message });
  }
});

// ─── PROTECTED: Info ─────────────────────────────────────────────────────
licenseRouter.get("/info", requireLicense, async (req, res) => {
  res.json({ ok: true, license: req.license, device: req.device });
});

// ─── PROTECTED: Heartbeat ────────────────────────────────────────────────
licenseRouter.post("/heartbeat", requireLicense, async (req, res) => {
  res.json({ ok: true, ttl: config.license.heartbeatTtlSec });
});

// ─── PROTECTED: List devices ─────────────────────────────────────────────
licenseRouter.get("/devices", requireLicense, async (req, res) => {
  try {
    const result = await licenseService.listDevices({
      licenseKey: req.licenseKey,
      currentDeviceId: req.deviceId,
    });
    if (!result.ok) {
      return res.status(result.code === "LICENSE_NOT_FOUND" ? 404 : 400).json({
        ok: false,
        code: result.code,
      });
    }
    res.json({
      ok: true,
      license: result.license,
      devices: result.devices,
    });
  } catch (err) {
    console.error("[License] /devices error:", err);
    res.status(500).json({ ok: false, code: "INTERNAL_ERROR", error: err.message });
  }
});

// ─── PROTECTED: Revoke device ────────────────────────────────────────────
licenseRouter.delete("/device/:deviceId", requireLicense, async (req, res) => {
  const targetDeviceId = req.params.deviceId;
  try {
    const result = await licenseService.revokeDevice({
      licenseKey: req.licenseKey,
      targetDeviceId,
    });
    if (!result.ok) {
      const status = result.code === "CANNOT_REVOKE_LAST_DEVICE" ? 409
                   : result.code === "DEVICE_NOT_FOUND" ? 404
                   : result.code === "LICENSE_NOT_FOUND" ? 404
                   : 400;
      return res.status(status).json({
        ok: false,
        code: result.code,
        error: mapCodeToVN(result.code),
      });
    }
    res.json({ ok: true, revoked: true });
  } catch (err) {
    console.error("[License] /device revoke error:", err);
    res.status(500).json({ ok: false, code: "INTERNAL_ERROR", error: err.message });
  }
});

// ─── Middleware: requireLicense ──────────────────────────────────────────
function requireLicense(req, res, next) {
  // Nếu LICENSE_REQUIRED=false (dev mode), skip
  if (!config.license.required) {
    req.license = { email: "dev@local", status: "active", maxDevices: 99 };
    req.device = { deviceId: "dev-device" };
    req.licenseKey = "dev-mode";
    req.deviceId = "dev-device";
    return next();
  }

  const { licenseKey, deviceId } = extractLicenseHeaders(req);

  if (!licenseKey || !deviceId) {
    return res.status(401).json({
      ok: false,
      code: "MISSING_LICENSE_HEADERS",
      error: "Thiếu X-License-Key hoặc X-Client-Id header.",
    });
  }

  // Validate async (cần gọi Supabase)
  licenseService
    .validate({
      licenseKey,
      deviceId,
      ip: getClientIp(req),
      ua: getUserAgent(req),
    })
    .then((result) => {
      if (!result.ok) {
        const status = result.code === "LICENSE_INACTIVE" ? 403
                     : result.code === "LICENSE_NOT_FOUND" ? 404
                     : result.code === "DEVICE_NOT_ENROLLED" ? 403
                     : 401;
        return res.status(status).json({
          ok: false,
          code: result.code,
          error: mapCodeToVN(result.code),
        });
      }
      req.license = result.license;
      req.device = result.device;
      req.licenseKey = licenseKey;
      req.deviceId = deviceId;
      next();
    })
    .catch((err) => {
      console.error("[License] requireLicense error:", err);
      res.status(500).json({ ok: false, code: "INTERNAL_ERROR", error: err.message });
    });
}

// ─── Helpers ─────────────────────────────────────────────────────────────
function parseDeviceNameFromUA(ua) {
  if (!ua) return "Unknown Device";
  // Đơn giản: "Word Desktop / Windows 11"
  const isWord = /Word/i.test(ua);
  const platform = /Windows/i.test(ua) ? "Windows"
                  : /Mac OS/i.test(ua) ? "macOS"
                  : /Linux/i.test(ua) ? "Linux"
                  : "Web";
  return isWord ? `Word Desktop / ${platform}` : `Web / ${platform}`;
}

function mapCodeToVN(code, detail) {
  const map = {
    LICENSE_NOT_FOUND: "License key không tồn tại. Vui lòng kiểm tra lại.",
    LICENSE_INACTIVE: `License đã bị ${detail?.status || "vô hiệu hóa"}. Liên hệ admin.`,
    DEVICE_LIMIT_EXCEEDED: `Đã đạt giới hạn ${detail?.active || "?"}/${detail?.max || "?"} thiết bị. Vào Supabase Studio để rotate.`,
    DEVICE_NOT_ENROLLED: "Thiết bị chưa được kích hoạt. Vui lòng nhập License Key.",
    INVALID_LICENSE_KEY: "License Key không hợp lệ (phải là UUID).",
    INVALID_EMAIL: "Email không hợp lệ.",
    INVALID_CLIENT_ID: "Client ID không hợp lệ.",
    CANNOT_REVOKE_LAST_DEVICE: "Không thể thu hồi thiết bị cuối cùng.",
    DEVICE_NOT_FOUND: "Thiết bị không tồn tại hoặc đã bị thu hồi.",
    RATE_LIMITED: "Quá nhiều yêu cầu. Vui lòng đợi 1 phút.",
    INTERNAL_ERROR: "Lỗi hệ thống. Vui lòng thử lại sau.",
  };
  return map[code] || code;
}

export { requireLicense };
