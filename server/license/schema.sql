-- ═══════════════════════════════════════════════════════════════════
-- AI Word Add-in — License Schema (Supabase / Postgres)
-- ═══════════════════════════════════════════════════════════════════
-- Hướng dẫn:
--   1. Tạo project tại https://supabase.com
--   2. Vào SQL Editor → New query
--   3. Copy toàn bộ file này → Run
--   4. Sau khi schema OK, copy URL + service_role key vào .env

-- ─── Extensions ────────────────────────────────────────────────────────────
create extension if not exists "pgcrypto";

-- ─── Tables ───────────────────────────────────────────────────────────────

create table if not exists public.licenses (
  id uuid primary key default gen_random_uuid(),
  license_key text not null unique,
  email text not null,
  status text not null default 'active'
    check (status in ('active', 'suspended', 'expired')),
  max_devices smallint not null default 1
    check (max_devices > 0 and max_devices <= 10),
  notes text,
  activated_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_licenses_email on public.licenses (email);
create index if not exists idx_licenses_status on public.licenses (status);

create table if not exists public.license_devices (
  id uuid primary key default gen_random_uuid(),
  license_id uuid not null references public.licenses(id) on delete cascade,
  device_id text not null,                    -- 32 hex chars (sha256 truncated)
  device_name text,                           -- Auto-detected: "Word Desktop / Windows 11"
  ip_address inet,
  user_agent text,
  first_seen_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  revoked_at timestamptz                      -- NULL = active
);

-- Mỗi license_key + device_id chỉ tồn tại 1 row active (revoked_at IS NULL)
create unique index if not exists uq_license_device_active
  on public.license_devices (license_id, device_id)
  where revoked_at is null;

create index if not exists idx_license_devices_license
  on public.license_devices (license_id);

create table if not exists public.license_events (
  id bigserial primary key,
  license_id uuid references public.licenses(id) on delete set null,
  device_id text,
  event text not null,                         -- 'activate' | 'heartbeat' | 'revoke' | 'reject' | 'info'
  detail jsonb,
  ip_address inet,
  user_agent text,
  created_at timestamptz not null default now()
);

create index if not exists idx_license_events_license
  on public.license_events (license_id, created_at desc);

-- ─── Triggers ──────────────────────────────────────────────────────────────

-- Tự update updated_at khi licenses thay đổi
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists trg_licenses_touch_updated on public.licenses;
create trigger trg_licenses_touch_updated
  before update on public.licenses
  for each row execute function public.touch_updated_at();

-- ─── Postgres Function: enroll_device ─────────────────────────────────────
-- Dùng trong RPC từ server để enroll device an toàn (chống race).
-- Trả JSON {ok, isNewDevice, code, ...}

create or replace function public.enroll_device(
  p_license_key text,
  p_device_id text,
  p_device_name text,
  p_ip inet,
  p_ua text
) returns jsonb
language plpgsql
security definer
as $$
declare
  v_license public.licenses%rowtype;
  v_existing public.license_devices%rowtype;
  v_active_count int;
begin
  -- Lock license row để tránh race condition
  select * into v_license
  from public.licenses
  where license_key = p_license_key
  for update;

  if not found then
    return jsonb_build_object('ok', false, 'code', 'LICENSE_NOT_FOUND');
  end if;

  if v_license.status != 'active' then
    return jsonb_build_object('ok', false, 'code', 'LICENSE_INACTIVE',
      'status', v_license.status);
  end if;

  -- Kiểm tra device đã tồn tại (chưa revoke)
  select * into v_existing
  from public.license_devices
  where license_id = v_license.id
    and device_id = p_device_id
    and revoked_at is null;

  if found then
    -- Touch last_seen_at
    update public.license_devices
    set last_seen_at = now()
    where id = v_existing.id;

    return jsonb_build_object(
      'ok', true,
      'isNewDevice', false,
      'licenseId', v_license.id,
      'deviceId', v_existing.id
    );
  end if;

  -- Đếm số devices active hiện tại
  select count(*) into v_active_count
  from public.license_devices
  where license_id = v_license.id
    and revoked_at is null;

  if v_active_count >= v_license.max_devices then
    return jsonb_build_object(
      'ok', false,
      'code', 'DEVICE_LIMIT_EXCEEDED',
      'active', v_active_count,
      'max', v_license.max_devices
    );
  end if;

  -- Insert device mới
  insert into public.license_devices
    (license_id, device_id, device_name, ip_address, user_agent)
  values
    (v_license.id, p_device_id, p_device_name, p_ip, p_ua);

  -- Touch activated_at lần đầu
  update public.licenses
  set activated_at = coalesce(activated_at, now())
  where id = v_license.id;

  -- Log event
  insert into public.license_events
    (license_id, device_id, event, detail, ip_address, user_agent)
  values
    (v_license.id, p_device_id, 'activate',
     jsonb_build_object('isNewDevice', true),
     p_ip, p_ua);

  return jsonb_build_object(
    'ok', true,
    'isNewDevice', true,
    'licenseId', v_license.id
  );
end $$;

-- ─── Function: validate_device ────────────────────────────────────────────
-- Trả JSON {ok, license, device, code}

create or replace function public.validate_device(
  p_license_key text,
  p_device_id text
) returns jsonb
language plpgsql
security definer
as $$
declare
  v_license public.licenses%rowtype;
  v_device public.license_devices%rowtype;
begin
  select * into v_license
  from public.licenses
  where license_key = p_license_key;

  if not found then
    return jsonb_build_object('ok', false, 'code', 'LICENSE_NOT_FOUND');
  end if;

  if v_license.status != 'active' then
    return jsonb_build_object('ok', false, 'code', 'LICENSE_INACTIVE',
      'status', v_license.status);
  end if;

  select * into v_device
  from public.license_devices
  where license_id = v_license.id
    and device_id = p_device_id
    and revoked_at is null;

  if not found then
    return jsonb_build_object('ok', false, 'code', 'DEVICE_NOT_ENROLLED');
  end if;

  -- Touch last_seen
  update public.license_devices
  set last_seen_at = now()
  where id = v_device.id;

  return jsonb_build_object(
    'ok', true,
    'license', jsonb_build_object(
      'id', v_license.id,
      'email', v_license.email,
      'status', v_license.status,
      'maxDevices', v_license.max_devices,
      'activatedAt', v_license.activated_at
    ),
    'device', jsonb_build_object(
      'id', v_device.id,
      'deviceId', v_device.device_id,
      'deviceName', v_device.device_name,
      'firstSeenAt', v_device.first_seen_at,
      'lastSeenAt', v_device.last_seen_at
    )
  );
end $$;

-- ─── RLS (Row Level Security) ──────────────────────────────────────────────
-- Server dùng service_role key nên bypass RLS.
-- Bật RLS phòng trường hợp ai đó dùng anon key truy cập.

alter table public.licenses enable row level security;
alter table public.license_devices enable row level security;
alter table public.license_events enable row level security;

-- Policy: chỉ service_role mới đọc/ghi (full quyền bypass qua service_role key)
-- Không cần tạo policy cụ thể vì RLS ON + không có policy = chặn mọi anon request.
-- service_role key sẽ bypass RLS hoàn toàn.

-- ─── Sample data (optional, comment out nếu không cần) ──────────────────
-- insert into public.licenses (license_key, email, status, max_devices, notes)
-- values
--   ('a1b2c3d4-e5f6-7890-abcd-ef1234567890', 'test@example.com', 'active', 1, 'Test license'),
--   ('f1e2d3c4-b5a6-7890-1234-567890abcdef', 'trial@example.com', 'active', 2, 'Trial 2 devices');
