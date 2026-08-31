/**
 * Device Fingerprint
 * Tạo device-id ổn định từ IP + User-Agent + Client UUID.
 * Hash sha256, cắt 32 hex chars.
 *
 * Lưu ý: device-id chỉ mang tính tương đối (client có thể spoof),
 * nhưng kết hợp với license_key + email + IP, đủ để chống share bậy
 * trong bối cảnh Word Add-in (cần license_key do admin cấp).
 */

import crypto from "crypto";

export function getClientIp(req) {
  // Ưu tiên X-Forwarded-For (khi chạy sau proxy/ngrok)
  const xff = req.headers["x-forwarded-for"];
  if (xff) return xff.split(",")[0].trim();
  return req.socket?.remoteAddress || req.ip || "0.0.0.0";
}

export function getUserAgent(req) {
  return req.headers["user-agent"] || "unknown";
}

export function buildDeviceId({ ip, userAgent, clientId }) {
  const data = [ip || "0.0.0.0", userAgent || "unknown", clientId || ""].join("|");
  return crypto.createHash("sha256").update(data).digest("hex").slice(0, 32);
}

export function deviceIdFromRequest(req) {
  const clientId = req.headers["x-client-id"] || req.headers["x-device-id"] || "";
  return buildDeviceId({
    ip: getClientIp(req),
    userAgent: getUserAgent(req),
    clientId,
  });
}
