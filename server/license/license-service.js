/**
 * License Service
 * Business logic cho license: activate, validate, list devices, revoke.
 * Tương tác với Supabase thông qua RPC functions (enroll_device, validate_device).
 */

import { getSupabase, isSupabaseConfigured } from "./supabase-client.js";
import { deviceIdFromRequest, getClientIp, getUserAgent } from "./device-fingerprint.js";

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export class LicenseService {
  constructor() {
    this._supabase = null;
  }

  _client() {
    if (!isSupabaseConfigured()) {
      throw new Error("Supabase chưa cấu hình (SUPABASE_URL / SUPABASE_SERVICE_KEY)");
    }
    if (!this._supabase) this._supabase = getSupabase();
    return this._supabase;
  }

  /**
   * Validate input format
   */
  validateInput({ email, licenseKey }) {
    if (!email || typeof email !== "string" || !EMAIL_REGEX.test(email)) {
      return { ok: false, code: "INVALID_EMAIL", httpStatus: 400 };
    }
    if (!licenseKey || typeof licenseKey !== "string" || !UUID_REGEX.test(licenseKey)) {
      return { ok: false, code: "INVALID_LICENSE_KEY", httpStatus: 400 };
    }
    return { ok: true };
  }

  /**
   * Activate (enroll) một device cho license.
   * Idempotent: nếu device đã active → touch last_seen.
   * Nếu license hết slot → trả DEVICE_LIMIT_EXCEEDED.
   */
  async activate({ email, licenseKey, clientId, deviceName, ip, ua }) {
    const inputCheck = this.validateInput({ email, licenseKey });
    if (!inputCheck.ok) return inputCheck;

    const deviceId = (clientId || "").toLowerCase();
    if (!deviceId || deviceId.length < 8) {
      return { ok: false, code: "INVALID_CLIENT_ID", httpStatus: 400 };
    }

    const sb = this._client();

    // Gọi RPC enroll_device (chạy atomic trong transaction)
    const { data, error } = await sb.rpc("enroll_device", {
      p_license_key: licenseKey.toLowerCase(),
      p_device_id: deviceId,
      p_device_name: deviceName || null,
      p_ip: ip || null,
      p_ua: ua || null,
    });

    if (error) {
      console.error("[License] RPC enroll_device error:", error);
      return { ok: false, code: "INTERNAL_ERROR", error: error.message, httpStatus: 500 };
    }

    if (!data?.ok) {
      const code = data?.code || "UNKNOWN";
      return {
        ok: false,
        code,
        detail: { active: data?.active, max: data?.max, status: data?.status },
        httpStatus: code === "DEVICE_LIMIT_EXCEEDED" ? 409
                  : code === "LICENSE_INACTIVE" ? 403
                  : code === "LICENSE_NOT_FOUND" ? 404
                  : 400,
      };
    }

    // Lấy thêm thông tin license + device để trả về client
    const { data: licenseRow } = await sb
      .from("licenses")
      .select("id, email, status, max_devices, activated_at")
      .eq("license_key", licenseKey.toLowerCase())
      .single();

    return {
      ok: true,
      isNewDevice: data.isNewDevice,
      license: {
        email: licenseRow?.email || email,
        status: licenseRow?.status,
        maxDevices: licenseRow?.max_devices,
        activatedAt: licenseRow?.activated_at,
      },
      device: {
        deviceId,
        firstSeenAt: data.isNewDevice ? new Date().toISOString() : null,
      },
    };
  }

  /**
   * Validate license + device cho mỗi request đến protected routes.
   * Trả {ok, license, device} hoặc {ok: false, code}.
   */
  async validate({ licenseKey, deviceId, ip, ua }) {
    if (!licenseKey || !UUID_REGEX.test(licenseKey)) {
      return { ok: false, code: "INVALID_LICENSE_KEY" };
    }
    if (!deviceId || deviceId.length < 8) {
      return { ok: false, code: "INVALID_DEVICE_ID" };
    }

    const sb = this._client();
    const { data, error } = await sb.rpc("validate_device", {
      p_license_key: licenseKey.toLowerCase(),
      p_device_id: deviceId.toLowerCase(),
    });

    if (error) {
      console.error("[License] RPC validate_device error:", error);
      return { ok: false, code: "INTERNAL_ERROR" };
    }

    if (!data?.ok) {
      return { ok: false, code: data?.code || "UNKNOWN" };
    }

    // Log heartbeat event (không await, fire-and-forget)
    sb.from("license_events").insert({
      license_id: data.license.id,
      device_id: deviceId,
      event: "heartbeat",
      ip_address: ip || null,
      user_agent: ua || null,
    }).then(() => {}).catch((e) => console.error("[License] Log event failed:", e));

    return { ok: true, license: data.license, device: data.device };
  }

  /**
   * Liệt kê tất cả devices của license (active + revoked).
   */
  async listDevices({ licenseKey }) {
    if (!licenseKey || !UUID_REGEX.test(licenseKey)) {
      return { ok: false, code: "INVALID_LICENSE_KEY" };
    }

    const sb = this._client();
    const { data: licenseRow, error: licError } = await sb
      .from("licenses")
      .select("id, email, status, max_devices, activated_at")
      .eq("license_key", licenseKey.toLowerCase())
      .single();

    if (licError || !licenseRow) {
      return { ok: false, code: "LICENSE_NOT_FOUND" };
    }

    const { data: devices, error: devError } = await sb
      .from("license_devices")
      .select("id, device_id, device_name, ip_address, user_agent, first_seen_at, last_seen_at, revoked_at")
      .eq("license_id", licenseRow.id)
      .order("first_seen_at", { ascending: false });

    if (devError) {
      console.error("[License] listDevices error:", devError);
      return { ok: false, code: "INTERNAL_ERROR" };
    }

    return {
      ok: true,
      license: {
        email: licenseRow.email,
        status: licenseRow.status,
        maxDevices: licenseRow.max_devices,
        activatedAt: licenseRow.activated_at,
      },
      devices: devices.map((d) => ({
        id: d.id,
        deviceId: d.device_id,
        deviceName: d.device_name,
        ip: d.ip_address,
        userAgent: d.user_agent,
        firstSeenAt: d.first_seen_at,
        lastSeenAt: d.last_seen_at,
        revokedAt: d.revoked_at,
        isCurrent: d.device_id === arguments[0]?.currentDeviceId,
      })),
    };
  }

  /**
   * Revoke một device.
   * Không cho phép revoke device cuối cùng (để tránh mất quyền truy cập).
   */
  async revokeDevice({ licenseKey, targetDeviceId }) {
    if (!licenseKey || !UUID_REGEX.test(licenseKey)) {
      return { ok: false, code: "INVALID_LICENSE_KEY" };
    }
    if (!targetDeviceId) {
      return { ok: false, code: "INVALID_DEVICE_ID" };
    }

    const sb = this._client();

    // Tìm license
    const { data: licenseRow } = await sb
      .from("licenses")
      .select("id")
      .eq("license_key", licenseKey.toLowerCase())
      .single();

    if (!licenseRow) {
      return { ok: false, code: "LICENSE_NOT_FOUND" };
    }

    // Đếm active devices
    const { count: activeCount } = await sb
      .from("license_devices")
      .select("id", { count: "exact", head: true })
      .eq("license_id", licenseRow.id)
      .is("revoked_at", null);

    if (activeCount !== null && activeCount <= 1) {
      return { ok: false, code: "CANNOT_REVOKE_LAST_DEVICE" };
    }

    // Revoke
    const { data, error } = await sb
      .from("license_devices")
      .update({ revoked_at: new Date().toISOString() })
      .eq("license_id", licenseRow.id)
      .eq("device_id", targetDeviceId.toLowerCase())
      .is("revoked_at", null)
      .select();

    if (error) {
      console.error("[License] revokeDevice error:", error);
      return { ok: false, code: "INTERNAL_ERROR" };
    }

    if (!data || data.length === 0) {
      return { ok: false, code: "DEVICE_NOT_FOUND" };
    }

    // Log event
    sb.from("license_events").insert({
      license_id: licenseRow.id,
      device_id: targetDeviceId,
      event: "revoke",
    }).then(() => {}).catch(() => {});

    return { ok: true, revoked: true };
  }

  /**
   * Heartbeat — chỉ touch last_seen.
   */
  async heartbeat({ licenseKey, deviceId }) {
    return this.validate({ licenseKey, deviceId });
  }
}

export const licenseService = new LicenseService();
