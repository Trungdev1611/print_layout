# Handoff — Create Sheet Set (phương án B)

Tài liệu này gồm 3 phần:

- **Phần A** — chốt yêu cầu + prompt để Cursor máy khác làm lại từ đầu.
- **Phần B** — code đã triển khai (tham chiếu).
- **Phần C** — khung hướng dẫn sử dụng cho người vẽ CAD.

Liên quan: `docs/HANDOFF_DRAWING_NAME.md` (tính năng tên bản vẽ) và
`docs/SHEET_SET_FIELD_SETUP.md` (cách gắn Field trong khung tên).

---

## Phần A — Chốt yêu cầu & prompt

### A.1 Bối cảnh đã thống nhất

Plugin có luồng: `PLSTT` (đánh số khung trong Model) → `PLAYOUT` (tạo layout cho
từng khung) → `PLPRINT` (xuất PDF). Cần thêm bước quản lý **Sheet Set** sau `PLAYOUT`.

Ba lớp dữ liệu, KHÔNG trộn:

| Vai trò | Nguồn | Ghi ở đâu |
| --- | --- | --- |
| STT = số hiệu bản vẽ = tên layout | `INNO-STT` (3 mode Simple/Advanced/Import của PLSTT) | Attribute khung trong Model |
| Tên bản vẽ | `INNO_NAME_DRAWING` (rule/Import, tùy chọn) | Attribute khung trong Model |
| Hiển thị trên khung tên | Field Sheet Set (người dùng tự gắn) | Attribute khung tên trong Paper Space |

### A.2 Quyết định chốt (phương án B)

1. **B, không phải A**: KHÔNG tạo attribute "số hiệu" thứ ba. STT kiêm luôn số hiệu và
   tên layout.
2. Sheet Set là bước riêng, chỉ xử lý **layout của DWG đang mở** (1 DWG). Kiến trúc để
   mở cho nhiều DWG sau này nhưng chưa làm.
3. UI WinForms dạng bảng, sắp xếp bằng **Move Up / Move Down** — KHÔNG kéo thả.
4. Có **Renumber 1…N** (chỉ đổi Sheet Number trên bảng + DST).
5. Có nút **Create / Update DST** tạo/ghi đè file `.dst`.
6. Có nút **Export PDF** ngay trong dialog, tận dụng lại luồng `PLPRINT`/`LayoutPlotter`
   hiện có, giữ đúng thứ tự bảng.
7. Seed dữ liệu: `Sheet Number` = tên layout (chính là STT); `Sheet Title` =
   `INNO_NAME_DRAWING` (thiếu thì bằng tên layout).
8. **Không** rename layout khi đổi Sheet Number. **Không** ghi ngược DST → attribute
   khung Model. Không đồng bộ hai chiều.
9. Việc "đổi trong sheet set thì khung tên tự đổi" giải quyết bằng **Field** trong khung
   tên (người dùng gắn tay), không phải code plugin.
10. Sau khi tạo DST, gọi `UpdateInMemoryDwgHints()` + `REGEN` để Field lấy được ngữ cảnh
    sheet set ngay.

### A.3 Prompt dán cho Cursor

```text
Dự án PrintLayoutAddin — plugin AutoCAD .NET (C#), đa target net48 (AutoCAD 2018-2024)
và net8.0-windows (AutoCAD 2025+). Đã có PLSTT (ghi INNO-STT + INNO_NAME_DRAWING lên
khung trong Model), PLAYOUT (mỗi khung -> 1 layout, tên layout = STT), PLPRINT
(LayoutPlotter xuất PDF qua PublishDsd).

Thêm tính năng Create Sheet Set theo "phương án B":
- STT kiêm số hiệu bản vẽ và tên layout; KHÔNG tạo attribute số hiệu thứ ba.
- Chỉ xử lý layout của DWG đang mở, nhưng viết model/dịch vụ mở cho nhiều DWG sau này.

Yêu cầu cụ thể:
1. Command mới PLSHEETSET + nút ribbon "Create Sheet Set" (đặt cạnh Build Layouts /
   Print), thêm vào danh sách ShortcutManager.
2. Dialog WinForms bảng: cột Use (checkbox), Order, Sheet Number, Drawing Name/Sheet
   Title, Layout, DWG. Seed: Sheet Number = tên layout; Sheet Title = INNO_NAME_DRAWING
   đọc từ khung Model theo STT, thiếu thì = tên layout.
3. Nút Move Up / Move Down (KHÔNG kéo thả), Renumber 1…N (chỉ đổi Sheet Number trên
   bảng + DST), Select All / None.
4. Nút Create / Update DST: tạo/ghi đè file .dst bằng Sheet Set COM API
   (AcSmSheetSetMgr). Dùng late binding qua Type.GetTypeFromProgID("AcSmComponents.<Class>.<ver>")
   để 1 bản build chạy nhiều đời AutoCAD, không tham chiếu interop cứng. Import layout
   bằng AcSmAcDbLayoutReference (InitNew/SetFileName/SetName) -> ImportSheet ->
   SetNumber/SetTitle -> InsertComponent. Lock/Unlock database, giải phóng COM ở finally.
   Sau khi thêm sheet, gọi UpdateInMemoryDwgHints() rồi REGEN document.
5. Nút Export PDF: tái dùng luồng PLPRINT. Ghi thứ tự + lựa chọn ra registry
   (PrintLayoutOrder/PrintLayoutChecked) rồi gọi _PLPRINT; PrintOptionsDialog thêm cờ
   preserveInputOrder để mở đúng thứ tự bảng.
6. KHÔNG rename layout, KHÔNG ghi ngược DST -> attribute Model.
7. Hiển thị hint trong dialog: dùng Field CurrentSheetNumber / CurrentSheetTitle trên
   khung tên; sau Create/Update DST cần lưu DWG và REGEN.

File dự kiến: Core/SheetSetService.cs (mới), UI/SheetSetDialog.cs (mới), Commands.cs,
UI/RibbonBuilder.cs, Core/ShortcutManager.cs, Core/FrameScanner.cs (đọc drawing name theo
STT), UI/PrintOptionsDialog.cs (cờ preserveInputOrder).

Build kiểm tra: dotnet build PrintLayoutAddin/PrintLayoutAddin.csproj -c Release
-f net8.0-windows (target net48 cần máy có SDK AutoCAD).
```

---

## Phần B — Code đã triển khai (tham chiếu)

### B.1 File thêm mới

- `PrintLayoutAddin/Core/SheetSetService.cs` — tạo `.dst` qua COM (late binding).
- `PrintLayoutAddin/UI/SheetSetDialog.cs` — dialog bảng quản lý sheet.

### B.2 `SheetSetService` — tạo DST bằng late binding COM

Ý chính: không tham chiếu `AcSmComponents<ver>Lib` (khác nhau theo đời AutoCAD). Thay vào
đó tạo COM object qua ProgID và gọi method bằng reflection.

```csharp
public class SheetSetEntry
{
    public bool Include { get; set; } = true;
    public int Order { get; set; }
    public string SheetNumber { get; set; }
    public string Title { get; set; }
    public string DwgPath { get; set; }
    public PrintableLayout Layout { get; set; }
    public string LayoutName => Layout?.Name ?? "";
    public string DwgName => Path.GetFileName(DwgPath ?? "");
}

public static void CreateOrReplace(string dstPath, string sheetSetName, IList<SheetSetEntry> entries)
{
    // ... validate + tạo thư mục
    object manager = CreateComObject("AcSmSheetSetMgr");
    object database = Invoke(manager, "CreateDatabase", dstPath, "", true);
    Invoke(database, "LockDb", database);
    var sheetSet = Invoke(database, "GetSheetSet");
    Invoke(sheetSet, "SetName", sheetSetName);

    foreach (var entry in entries)
    {
        var layoutRef = CreateComObject("AcSmAcDbLayoutReference");
        Invoke(layoutRef, "InitNew", sheetSet);
        Invoke(layoutRef, "SetFileName", entry.DwgPath);
        Invoke(layoutRef, "SetName", entry.Layout.Name);

        var sheet = Invoke(sheetSet, "ImportSheet", layoutRef);
        Invoke(sheet, "SetNumber", entry.SheetNumber ?? "");
        Invoke(sheet, "SetTitle", entry.Title);
        Invoke(sheetSet, "InsertComponent", sheet, null);
    }

    // Field CurrentSheetNumber/Title cần hint này trong DWG đang mở.
    try { Invoke(database, "UpdateInMemoryDwgHints"); } catch { }
    // finally: UnlockDb(database, true); Close(database); FinalReleaseComObject(...)
}
```

Tạo COM theo ProgID có version, fallback không version:

```csharp
private static object CreateComObject(string className)
{
    int major = GetAcadMajorVersion(); // đọc ACADVER, lấy phần major
    foreach (var progId in new[] { $"AcSmComponents.{className}.{major}", $"AcSmComponents.{className}" })
    {
        var type = Type.GetTypeFromProgID(progId, false);
        if (type != null) return Activator.CreateInstance(type);
    }
    throw new InvalidOperationException($"Sheet Set component '{className}' unavailable.");
}

private static object Invoke(object target, string method, params object[] args)
    => target.GetType().InvokeMember(method,
        System.Reflection.BindingFlags.InvokeMethod, null, target, args,
        System.Globalization.CultureInfo.InvariantCulture);
```

Lưu ý quan trọng:
- Sheet Set COM API chỉ chạy **trong tiến trình AutoCAD**, không chạy exe rời.
- Phải `LockDb` trước khi sửa, `UnlockDb(db, true)` để commit, và giải phóng COM ở
  `finally` để tránh khóa file `.dst`.

### B.3 `FrameScanner` — đọc tên bản vẽ theo STT

```csharp
public static string ReadDrawingName(BlockReference br, Config cfg)
    => ReadAttribute(br, cfg.DrawingNameTag);

// Trả về map STT -> DrawingName, để dialog seed cột Sheet Title.
public static Dictionary<string, string> CollectDrawingNamesByStt(Database db) { /* duyệt ModelSpace */ }
```

### B.4 `Commands.PLSHEETSET`

```csharp
[CommandMethod("PLSHEETSET")]
public void PlSheetSet()
{
    // ... license + EnsureSavedForPublish
    var layouts = LayoutPlotter.GetPrintableLayouts(doc.Database)
        .Where(x => !string.Equals(x.Name, Config.Instance.TemplateLayout, StringComparison.OrdinalIgnoreCase))
        .ToList();
    var drawingNames = FrameScanner.CollectDrawingNamesByStt(doc.Database);
    var defaultDstPath = Path.ChangeExtension(doc.Name, ".dst");

    using (var dlg = new SheetSetDialog(layouts, drawingNames, doc.Name, defaultDstPath))
    {
        if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
        if (dlg.ExportLayouts == null || dlg.ExportLayouts.Count == 0) return;
        SavePrintLayoutSelection(dlg.ExportLayouts); // ghi order+checked ra registry
    }
    // Export PDF -> tái dùng PLPRINT, mở đúng thứ tự bảng
    doc.SendStringToExecute("_PLPRINT ", true, false, true);
}
```

`SavePrintLayoutSelection` ghi `PrintLayoutOrder` / `PrintLayoutChecked`
(MultiString) vào `HKCU\Software\PrintLayoutAddin`.

### B.5 `SheetSetDialog` (WinForms)

- `BindingList<SheetSetEntry>` + `DataGridView` (AutoGenerateColumns=false).
- Cột: Use / Order / Sheet Number / Drawing Name-Title / Layout / DWG.
- `MoveSelected(±1)`: remove + insert trong BindingList rồi `RefreshOrders`.
- `Renumber()`: gán lại Sheet Number 1..N cho các dòng Include.
- `CreateDst()`: gọi `SheetSetService.CreateOrReplace`, sau đó `Editor.Regen()`.
- `RequestExport()`: set `ExportLayouts` theo thứ tự Include rồi `DialogResult.OK`.
- Header hiện 2 hint: quy tắc seed (workflow B) và cách dùng Field khung tên.

### B.6 `PrintOptionsDialog` — cờ giữ thứ tự

```csharp
public PrintOptionsDialog(Database db, IEnumerable<PrintableLayout> layouts,
    string templateLayoutName, string defaultPdfPath, bool preserveInputOrder = false)
// preserveInputOrder = true -> dùng thứ tự truyền vào thay vì thứ tự lưu registry.
```

### B.7 Ribbon & Shortcut

- `RibbonBuilder`: thêm `btnSheetSet` (Create Sheet Set) trước nút Print,
  icon `DrawSheetSetIcon`.
- `ShortcutManager.Defs`: thêm `new ShortcutDef("PLSHEETSET", "Create Sheet Set", Keys.None)`.

### B.8 Build

```bash
dotnet build PrintLayoutAddin/PrintLayoutAddin.csproj -c Release -f net8.0-windows
```
Đã pass 0 warning/0 error trên target net8. Target net48 cần máy có SDK AutoCAD
(`/p:AutoCADPath=...`). Phần tạo `.dst` phải test thực tế trong AutoCAD vì COM Sheet Set
chỉ sống trong tiến trình AutoCAD.

---

## Phần C — Hướng dẫn sử dụng (nháp cho người vẽ CAD)

> Phần này viết để sau chuyển thành tài liệu hướng dẫn chính thức.

### C.1 Chuẩn bị một lần

1. Sửa **block khung tên** (title block) dùng trong layout mẫu:
   - Ô số bản vẽ: chèn Field **CurrentSheetNumber**.
   - Ô tên bản vẽ: chèn Field **CurrentSheetTitle**.
   - Lưu và reload xref.
2. Không dùng attribute `INNO-STT` / `INNO_NAME_DRAWING` của khung Model làm chữ trong
   khung tên — hai chỗ đó khác nhau.

### C.2 Quy trình từng bước

1. **PLSTT** — chọn polyline dẫn đường, chọn khung, chọn chế độ đánh số
   (Simple / Advanced / Import). Nếu cần, bật gán tên bản vẽ. Kết quả: mỗi khung có STT
   (và tên) riêng.
2. **PLAYOUT** — sinh layout cho từng khung; tên layout = STT.
3. **PLSHEETSET** — mở bảng Create Sheet Set:
   - Cột **Sheet Number** mặc định = STT; **Sheet Title** = tên bản vẽ.
   - Dùng **Move Up / Down** để sắp thứ tự in.
   - **Renumber 1…N** nếu muốn đánh lại số tờ liên tục.
   - Chọn nơi lưu file **.dst** → bấm **Create / Update DST**.
   - (Tùy chọn) **Export PDF** để in luôn theo thứ tự bảng.
4. **Lưu DWG** và chạy **REGEN** nếu khung tên chưa cập nhật Field.

### C.3 Câu hỏi thường gặp

- **Sửa số/tên trong bảng có đổi tên layout không?** Không. Chỉ đổi trong bảng và file
  `.dst`.
- **Sửa trong Sheet Set Manager của AutoCAD có đổi ngược lại attribute khung Model
  không?** Không. Muốn đổi STT/tên trên khung thì chạy lại `PLSTT`.
- **Vì sao khung tên tự hiện đúng số/tên?** Vì nó dùng Field trỏ vào Sheet Set, không
  phải chữ cứng. Đổi `.dst` + REGEN là khung tên đổi theo.
- **Đổi số/tên trong bảng mà chưa bấm Create/Update DST?** File `.dst` chưa đổi, khung
  tên chưa cập nhật.

### C.4 Giới hạn hiện tại

- Sheet Set chỉ gom layout của **1 DWG đang mở**.
- Không tự rename layout theo Sheet Number.
- Không đồng bộ hai chiều giữa `.dst` và attribute khung Model.
