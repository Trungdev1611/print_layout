# Hướng dẫn gửi cài đặt & cấp License

Tài liệu gửi nội bộ (bạn / admin) và có thể copy phần **A** / **C** cho đội CAD.

---

## Tóm tắt luồng

```
1. Bạn gửi file cài → đội CAD cài add-in
2. Đội CAD mở AutoCAD → lấy Machine ID (PLLICENSE) → gửi lại cho bạn
3. Bạn dùng PLLicenseGen (nội bộ) → cấp key theo Machine ID + ngày hết hạn
4. Đội CAD dán key vào PLLICENSE → Activate
```

Key **gắn với 1 máy**. Máy khác = Machine ID khác → cần key mới.

---

## A. Gửi gì cho đội CAD (cài đặt)

Gửi **một** trong hai file (đã build sẵn trong `installer\`):

| File | Cách dùng |
| --- | --- |
| **`PrintLayoutAddin-Setup.zip`** (khuyến nghị) | Giải nén → chạy `install.bat` |
| **`PrintLayoutAddinSetup.exe`** | Double-click để cài |

**Không gửi:**

- Source code / cả repo
- **`PLLicenseGen.exe`** (công cụ cấp key — chỉ giữ nội bộ)
- DLL lẻ ngoài package trên

### A.1 Việc đội CAD làm sau khi nhận file

1. Cài bằng zip (`install.bat`) hoặc `.exe`.
2. **Đóng hết AutoCAD** rồi mở lại (add-in load khi start).
3. Gõ lệnh thử (ví dụ `PLSTT` / `PLSHEETSET`). Nếu chưa có license → hộp thoại kích hoạt hiện ra, hoặc gõ:

   ```
   PLLICENSE
   ```

4. Trong dialog **License**:
   - Copy **Machine ID** (nút Copy).
   - Gửi Machine ID về cho người cấp license (kèm tên máy / người dùng nếu cần).
5. Nhận key → dán vào ô **License Key** → **Activate** → **Close**.
6. Chạy lại lệnh plugin để xác nhận đã vào được.

---

## B. Việc bạn làm để lấy / cấp key (nội bộ)

### B.1 Chuẩn bị

- Tool: `license-generator\bin\Release\net48\PLLicenseGen.exe`  
  (build: `dotnet build license-generator\LicenseGenerator.csproj -c Release`)
- **Không** đưa file này cho đội CAD / khách.
- Cần từ đội CAD: **Machine ID** (chuỗi copy từ `PLLICENSE`).

### B.2 Cấp 1 key (GUI)

1. Chạy `PLLicenseGen.exe`.
2. Tab / chế độ **cấp 1 key**:
   - Dán **Machine ID**
   - Chọn / nhập **ngày hết hạn** (`yyyy-MM-dd`)
   - (Tuỳ chọn) **Note** — tên user / phòng ban (hiện trong trạng thái license)
3. Tạo key → copy chuỗi `PLA1-...`
4. Gửi key cho đúng người đã gửi Machine ID đó.

### B.3 Cấp nhiều key (CSV)

1. Trong `PLLicenseGen`: xuất **file mẫu** CSV.
2. Điền cột (theo mẫu): `user`, `machine_id`, `expire`, `note`.
3. **Nhập CSV** → tool sinh key hàng loạt.
4. Xuất CSV kết quả (có cột `license_key`) → gửi từng key cho đúng máy.

### B.4 CLI (tuỳ chọn)

```text
PLLicenseGen --mid <MachineId> --expire YYYY-MM-DD [--note "Tên NV"]
```

---

## C. Tin nhắn mẫu gửi đội CAD

Có thể copy-paste:

```text
Cài PrintLayoutAddin:

1) Giải nén PrintLayoutAddin-Setup.zip → chạy install.bat
   (hoặc chạy PrintLayoutAddinSetup.exe)
2) Đóng AutoCAD rồi mở lại
3) Gõ lệnh: PLLICENSE
4) Copy Machine ID → gửi lại cho mình (kèm tên máy / người dùng)
5) Mình gửi License Key → các bạn dán vào PLLICENSE → Activate

Lưu ý: Key chỉ dùng được trên đúng máy đã gửi Machine ID.
```

---

## D. Checklist nhanh

**Bạn gửi đi**

- [ ] `PrintLayoutAddin-Setup.zip` **hoặc** `PrintLayoutAddinSetup.exe`
- [ ] Hướng dẫn ngắn (phần C)
- [ ] **Không** gửi `PLLicenseGen`

**Đội CAD gửi lại**

- [ ] Machine ID (từ `PLLICENSE`)
- [ ] (Nên có) tên người / máy / ngày cần dùng đến

**Bạn gửi key**

- [ ] Key `PLA1-...` khớp đúng Machine ID
- [ ] Ngày hết hạn đã thống nhất

**Đội CAD hoàn tất**

- [ ] Activate thành công trong `PLLICENSE`
- [ ] Chạy được lệnh plugin (không còn bị chặn license)

---

## E. Lỗi thường gặp

| Triệu chứng | Nguyên nhân / xử lý |
| --- | --- |
| Lệnh plugin mở dialog license mãi | Chưa Activate, hoặc key sai |
| Wrong machine | Key của máy khác — cần Machine ID máy đang dùng |
| Expired | Hết hạn — cấp key mới với ngày mới |
| Cài xong không thấy lệnh | Chưa restart AutoCAD; hoặc cài nhầm user Windows |
| SmartScreen chặn `.exe` | Dùng bản `.zip` + `install.bat` |

---

## F. Vị trí file sau khi build (máy bạn)

| Thành phần | Đường dẫn |
| --- | --- |
| Zip cài (gửi CAD) | `installer\PrintLayoutAddin-Setup.zip` |
| Exe cài (gửi CAD) | `installer\PrintLayoutAddinSetup.exe` |
| Tool cấp key (nội bộ) | `license-generator\bin\Release\net48\PLLicenseGen.exe` |

Build lại package:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_portable_zip.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"
powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"
```
