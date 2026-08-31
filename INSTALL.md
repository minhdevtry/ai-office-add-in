# 📥 Hướng Dẫn Cài Đặt AI Word Add-in Lên Máy Khác

> **Server đã sẵn sàng tại:** `https://word-bridge.2tocom.space`
> Manifest đã được cập nhật, không cần build lại.

---

## 🎯 Tổng quan

- **Server backend**: chạy trên máy chủ (Ubuntu) + Cloudflare Tunnel → public URL `https://word-bridge.2tocom.space`
- **Mỗi máy Word** chỉ cần file `manifest.xml` để cài add-in
- **License** (Supabase) đã được setup trên server — bạn chỉ cần tạo license key cho mỗi khách hàng qua Supabase Studio

---

## ✅ Checklist trước khi bắt đầu (làm 1 lần trên server)

### 1. Setup Supabase (license DB) — NẾU CHƯA LÀM

1. Truy cập [https://poysgcsiwzhplkhqbgsw.supabase.co](https://poysgcsiwzhplkhqbgsw.supabase.co)
2. Vào **SQL Editor** → **New query**
3. Copy toàn bộ nội dung file `server/license/schema.sql` → paste vào → bấm **Run**
4. Tạo license cho từng khách hàng:
   ```sql
   INSERT INTO public.licenses (license_key, email, status, max_devices, notes)
   VALUES (
     gen_random_uuid()::text,  -- hoặc tự đặt UUID cụ thể
     'email_khach@example.com',
     'active',
     1,                        -- số máy cho phép (thường = 1)
     'Tên khách hàng / ghi chú'
   );
   ```
5. **Lưu lại `license_key`** (UUID) — sẽ gửi cho khách hàng để activate.

### 2. Kiểm tra server + tunnel

```bash
# SSH vào server, chạy:
curl -sS https://word-bridge.2tocom.space/api/status
# Phải trả về JSON với status: "ok"
```

Nếu OK → tiếp tục. Nếu không → check:
```bash
sudo systemctl status cloudflared-word-bridge
sudo systemctl status word-bridge  # service cho app server (xem phần dưới)
```

### 3. Đảm bảo server Node.js chạy persistent

Nếu chưa có systemd service cho Node app, tạo file `/etc/systemd/system/word-bridge.service`:

```ini
[Unit]
Description=AI Word Add-in Backend
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=lucas
WorkingDirectory=/home/lucas/Documents/code/ai-office-add-in
ExecStart=/usr/bin/node server/server.js
Restart=on-failure
RestartSec=5
EnvironmentFile=/home/lucas/Documents/code/ai-office-add-in/.env
LimitNOFILE=65536

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now word-bridge
sudo systemctl status word-bridge
```

---

## 🖥️ Cài Add-in Lên Máy Word (Windows / Mac)

### Bước 1: Copy file `manifest.xml`

Manifest đã có sẵn tại `https://word-bridge.2tocom.space/manifest.xml` — tải về máy:
- **Trực tiếp:** mở trình duyệt tại `https://word-bridge.2tocom.space/manifest.xml` → Ctrl+S để lưu
- **Hoặc từ repo:** copy file `manifest.xml` trong project này

Lưu vào 1 vị trí dễ nhớ, ví dụ:
- Windows: `C:\Users\<username>\Documents\AI-Word-Addin\manifest.xml`
- macOS: `~/Documents/AI-Word-Addin/manifest.xml`

### Bước 2: Sideload vào Word

#### 🅰️ Word 365 Web (đơn giản nhất)

1. Mở [https://word.office.com](https://word.office.com) → mở 1 file bất kỳ
2. Vào thẻ **Insert (Chèn)** trên thanh ribbon
3. Nhấp **Add-ins (Tiện ích bổ sung)** → **More Add-ins (Thêm tiện ích)**
4. Bấm tab **MY ADD-INS** ở trên cùng
5. Kéo xuống dưới → **Upload My Add-in (Tải tiện ích của tôi lên)**
6. Chọn **MANIFEST.XML FILE** (hoặc "Choose File" tùy phiên bản)
7. Browse đến file `manifest.xml` vừa lưu → **Install**
8. Quay lại ribbon → thẻ **Home** → tìm nút **Trợ lý AI** → click để mở sidebar

#### 🅱️ Word Desktop trên Windows

> **⚠️ Lưu ý:** "Trusted Add-in Catalogs" trong Trust Center chỉ dùng cho **shared network folders** (UNC paths), KHÔNG hoạt động với local folder thường. Dùng cách dưới đây.

**Cách 1 (Recommend): Cài qua file manifest cục bộ**

1. **Tải manifest về máy**: mở trình duyệt → truy cập `https://word-bridge.2tocom.space/manifest.xml` → Ctrl+S → lưu vào `C:\Users\<tên-user>\Documents\AI-Word-Addin\manifest.xml`
2. **Mở Windows Explorer** → tìm file `manifest.xml` vừa lưu
3. **Double-click** file `manifest.xml` → Word sẽ mở + hiện popup **"This add-in wants to make changes. Install?"** → bấm **Install**
4. Đóng Word hoàn toàn → mở lại Word
5. Ribbon **Home** → tìm nhóm **AI Assistant** ở **cuối ribbon** (bên phải) → nút **Trợ lý AI**

**Cách 2: Cài từ trang Word trong Office (Workaround khi không double-click được)**

1. Tải `manifest.xml` về máy (như bước 1 trên)
2. Mở Word → **File** → **Options** → **Trust Center** → **Trust Center Settings...**
3. Chọn **Trusted Add-in Catalogs** ở sidebar trái
4. Ở ô **Catalog URL**, paste đường dẫn dạng `file:///`:
   ```
   file:///C:/Users/YourName/Documents/AI-Word-Addin
   ```
5. Tích **Show in Menu** → OK → **khởi động lại Word**
6. Vào **Insert** → **My Add-ins** → chọn tab **Shared Folder** (hoặc **Developer Add-ins** tùy phiên bản)
7. Nếu thấy "Trợ lý AI Word" → click để thêm vào document
8. Ribbon **Home** → cuộn sang phải → nhóm **AI Assistant** → nút **Trợ lý AI**

**Cách 3 (Pro): Cài bằng PowerShell registry (chỉ dùng nếu 2 cách trên fail)**

Mở PowerShell **as Administrator** rồi chạy:
```powershell
$manifestPath = "C:\Users\YourName\Documents\AI-Word-Addin\manifest.xml"
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedAddins\Trợ\ lý\ AI\ Word"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "Url" -Value $manifestPath
Set-ItemProperty -Path $key -Name "Id" -Value "f47ac10b-58cc-4372-a567-0e02b2c3d479"
```
(Trong đó `Id` lấy từ `<Id>...</Id>` trong manifest)

Sau đó mở lại Word.

**Vị trí nút trên ribbon:**
- Nút **Trợ lý AI** nằm trong nhóm **AI Assistant** ở **tab Home**, vị trí **cuối ribbon** (bên phải, sau các nhóm Editing/Clipboard)
- Nếu ribbon bị thu nhỏ, click nút **`>`** cuối ribbon để mở rộng
- Hoặc kéo giãn cửa sổ Word rộng ra

#### 🅲 Word Desktop trên macOS

1. Tạo thư mục wef của Word (nếu chưa có):
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
2. Copy `manifest.xml` vào đó:
   ```bash
   cp ~/Downloads/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
3. **Khởi động lại Word**
4. Vào **Insert** → **My Add-ins** → **Developer Add-ins** → chọn **Trợ lý AI Word**

---

## 🔐 Kích hoạt License (LẦN ĐẦU)

Sau khi mở Add-in lần đầu tiên, modal kích hoạt sẽ hiện ra:

1. Nhập **Email** đã đăng ký license
2. Nhập **License Key** (UUID) mà admin cung cấp
3. Bấm **Kích hoạt**
4. Modal đóng → có thể bắt đầu dùng Add-in

> **Lưu ý:** License chỉ bind với 1 máy (1 device-id UUID được tạo tự động trong browser). Nếu muốn đổi máy, vào **Supabase Studio** → bảng `license_devices` → set `revoked_at = NOW()` cho máy cũ → kích hoạt lại trên máy mới.

---

## 🛠️ Cho người quản trị (admin)

### Tạo license mới

Vào Supabase Studio → SQL Editor:
```sql
INSERT INTO public.licenses (license_key, email, status, max_devices, notes)
VALUES (
  gen_random_uuid()::text,
  'customer@example.com',
  'active',
  1,
  'Khách hàng: Nguyễn Văn A, mua 2026-08-31'
)
RETURNING license_key, email;
```

Copy `license_key` (UUID) → gửi cho khách hàng qua email/Zalo kèm hướng dẫn ở trên.

### Xem danh sách devices đang active

```sql
SELECT
  l.email,
  ld.device_id,
  ld.device_name,
  ld.ip_address,
  ld.first_seen_at,
  ld.last_seen_at,
  CASE WHEN ld.revoked_at IS NULL THEN '✅ Active' ELSE '❌ Revoked' END AS status
FROM public.licenses l
LEFT JOIN public.license_devices ld ON ld.license_id = l.id
ORDER BY l.email, ld.first_seen_at DESC;
```

### Thu hồi thiết bị (để user đổi máy mới)

```sql
UPDATE public.license_devices
SET revoked_at = NOW()
WHERE device_id = 'DEVICE_ID_CU'  -- paste từ bảng trên
  AND revoked_at IS NULL
RETURNING id, device_id, revoked_at;
```

Sau đó bảo user mở lại Add-in trên máy mới → modal kích hoạt sẽ tự hiện ra.

### Vô hiệu hóa license (khách dừng sử dụng)

```sql
UPDATE public.licenses
SET status = 'suspended'
WHERE email = 'customer@example.com';
```

---

## 🔍 Kiểm tra nhanh

| Test | URL/Command | Kết quả mong đợi |
|---|---|---|
| Server online | `https://word-bridge.2tocom.space/api/status` | JSON với `status: "ok"` |
| UI load | Mở `https://word-bridge.2tocom.space/index.html` | Trang HTML giao diện |
| Public config | `curl https://word-bridge.2tocom.space/api/license/public-config` | JSON `{ok:true, aiModel, aiDefaultEffort}` |
| Tools catalog | `curl https://word-bridge.2tocom.space/tools` | Danh sách 14 tools |
| Tunnel status | `sudo systemctl status cloudflared-word-bridge` | active (running) |
| Server status | `sudo systemctl status word-bridge` | active (running) |

---

## 🆘 Troubleshooting

### Add-in không load / không kết nối được
1. Mở DevTools trong Taskpane (Word → View → Office Add-ins → My Add-ins → click add-in → ... menu → Inspect)
2. Tab **Network** → xem request tới `word-bridge.2tocom.space` có pass không
3. Nếu 401 → chưa activate license, mở modal lại và nhập key
4. Nếu connection error → server có thể down, check `sudo systemctl status word-bridge`

### "Thiết bị đã đạt giới hạn 1/1"
License đã được dùng ở máy khác. Vào Supabase → bảng `license_devices` → revoke máy cũ (xem trên).

### Modal kích hoạt không hiện
- Word đã cache manifest. Thoát Word hoàn toàn + clear Office cache:
  - Windows: `%localappdata%\Microsoft\Office\16.0\OfficeFileCache\`
  - macOS: `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
- Hoặc vào **Insert** → **My Add-ins** → tìm "Trợ lý AI Word" → click lại

### Server restart cần thiết
```bash
sudo systemctl restart word-bridge
sudo systemctl restart cloudflared-word-bridge  # nếu tunnel có vấn đề
```

---

## 📞 Liên hệ hỗ trợ

- Email: support@2tocom.space
- Cloudflare Tunnel: xem logs tại `sudo journalctl -u cloudflared-word-bridge -f`
- App server logs: `tail -f /tmp/word-bridge-server.log`
- Supabase: [https://poysgcsiwzhplkhqbgsw.supabase.co](https://poysgcsiwzhplkhqbgsw.supabase.co)
