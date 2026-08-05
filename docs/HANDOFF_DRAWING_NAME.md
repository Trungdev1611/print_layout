# Handoff — Gán tên bản vẽ song song với STT

Tài liệu này để dán lại cho Cursor trên máy khác làm lại tính năng từ đầu,
hoặc để review nhanh những gì đã thay đổi.

## Cách dùng ở máy công ty (repo khác, code giống)

1. Copy file này sang repo công ty, ví dụ đặt vào `docs/`.
2. Mở Cursor, kéo file vào chat rồi ra lệnh:

   > Đọc `docs/HANDOFF_DRAWING_NAME.md` và implement đúng theo phần 1 và phần 2.
   > Sau khi sửa xong build bằng
   > `dotnet build PrintLayoutAddin/PrintLayoutAddin.csproj -c Release -f net8.0-windows`
   > và báo lại các file đã thay đổi.

3. Nếu code bên đó đã lệch so với bản này, bảo Cursor bám theo mô tả yêu cầu ở phần 1
   và coi code phần 2 là tham chiếu, không copy máy móc.

---

## 1. Prompt để dán cho Cursor

```text
Dự án: PrintLayoutAddin — plugin AutoCAD .NET (C#), đa target net48 (AutoCAD 2018-2024)
và net8.0-windows (AutoCAD 2025+). Lệnh PLSTT đi theo polyline dẫn đường và ghi số thứ tự
vào attribute INNO-STT của các block khung trong ModelSpace.

Yêu cầu: cho phép gán THÊM "tên bản vẽ" vào một attribute thứ hai, song song với STT,
trong cùng một lần chạy PLSTT.

Ràng buộc bắt buộc:
1. Tên attribute tên bản vẽ phải khai báo tập trung một chỗ duy nhất, mặc định
   INNO_NAME_DRAWING, override được qua config.json bằng khóa "drawingNameTag".
   Sau này khách đổi tên attribute thì chỉ sửa đúng 1 chỗ.
2. Tab Import: đọc thêm cột DrawingName cùng hàng với FrameNumber. Chấp nhận các alias
   DrawingName, Drawing_Name, Name_Drawing, INNO_NAME_DRAWING, "Tên bản vẽ", tenbanve.
   Template export (CSV + XLSX) phải có sẵn cột DrawingName.
3. Có thêm rule sinh tên tự động giống Simple mode: Prefix, Suffix, Start, Step,
   Padding, Skip. Dùng lại NumberGenerator.GenerateSimple, không viết bộ sinh mới.
4. Có checkbox "Assign drawing names" để bật/tắt. Tắt thì hành vi giống hệt bản cũ.
5. Lưới Preview hiển thị 2 cột: Frame Number và Drawing Name.
6. Khi ghi: khung nào không có attribute tên thì BỎ QUA phần tên nhưng VẪN ghi STT
   bình thường, và đếm số khung bị thiếu để in ra command line.
7. Tên rỗng thì bỏ qua, không ghi đè attribute bằng chuỗi rỗng.
8. KHÔNG dùng XData fallback cho tên bản vẽ — fallback chỉ dành riêng cho STT như cũ.
9. PLAUTO khi tạo block native PL_ phải tạo sẵn CẢ HAI attribute definition.
10. Giữ nguyên overload cũ của ApplyNumbersToSelectedFrames để không phá code gọi cũ.
11. Cập nhật cả config.json trong PrintLayoutAddin/ lẫn trong
    PrintLayoutAddin.bundle/Contents/net48 và /net8.

Các file cần sửa: Core/Config.cs, Core/FrameScanner.cs, Core/SttAssigner.cs,
Core/NumberImporter.cs, Core/NativeFrameBuilder.cs, UI/SttOptionsDialog.cs, Commands.cs,
config.json.

Build kiểm tra bằng:
dotnet build PrintLayoutAddin/PrintLayoutAddin.csproj -c Release -f net8.0-windows
(target net48 cần máy có cài AutoCAD SDK, có thể bỏ qua nếu máy không có).
```

---

## 2. Code mẫu các điểm chính

### 2.1 `Core/Config.cs` — nơi khai báo tag duy nhất

```csharp
public class Config
{
    // Central defaults for frame attributes. Deployments can override either
    // value in config.json without recompiling the add-in.
    public const string DefaultFrameNumberTag = "INNO-STT";
    public const string DefaultDrawingNameTag = "INNO_NAME_DRAWING";

    public string AttributeTag { get; set; } = DefaultFrameNumberTag;
    public string DrawingNameTag { get; set; } = DefaultDrawingNameTag;
    // ...

    private static Config Load()
    {
        // ...
        cfg.AttributeTag    = ExtractString(json, "attributeTag")    ?? cfg.AttributeTag;
        cfg.DrawingNameTag  = ExtractString(json, "drawingNameTag")  ?? cfg.DrawingNameTag;
        // ...
    }
}
```

`config.json`:

```json
{
  "attributeTag": "INNO-STT",
  "drawingNameTag": "INNO_NAME_DRAWING",
  "vpLayer": "360D-Mview",
  "xdataAppName": "PLADDIN_STT",
  "templateLayout": "Layout1"
}
```

### 2.2 `Core/FrameScanner.cs` — tách hàm ghi attribute dùng chung

```csharp
public static bool WriteStt(BlockReference br, string value, Config cfg, Transaction tr)
{
    if (WriteAttribute(br, cfg.AttributeTag, value, tr)) return true;

    // STT keeps its legacy XData fallback because downstream layout
    // generation can read it when an xref does not expose attributes.
    EnsureRegApp(br.Database, cfg.XdataAppName, tr);
    var brW = (BlockReference)tr.GetObject(br.ObjectId, OpenMode.ForWrite);
    brW.XData = new ResultBuffer(
        new TypedValue((int)DxfCode.ExtendedDataRegAppName, cfg.XdataAppName),
        new TypedValue((int)DxfCode.ExtendedDataAsciiString, value ?? ""));
    return true;
}

public static bool WriteDrawingName(BlockReference br, string value, Config cfg, Transaction tr)
{
    if (string.IsNullOrWhiteSpace(value)) return false;
    return WriteAttribute(br, cfg.DrawingNameTag, value, tr);
}

private static bool WriteAttribute(BlockReference br, string tag, string value, Transaction tr)
{
    if (string.IsNullOrWhiteSpace(tag)) return false;
    foreach (ObjectId attId in br.AttributeCollection)
    {
        var att = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
        if (att == null) continue;
        if (string.Equals(att.Tag, tag, StringComparison.OrdinalIgnoreCase))
        {
            att.TextString = value;
            return true;
        }
    }
    return false;
}
```

### 2.3 `Core/SttAssigner.cs` — ghi hai giá trị trong cùng transaction

```csharp
public class SttAssignResult
{
    // ... các field cũ
    public int DrawingNamesAssigned;
    public int DrawingNameAttributesMissing;
}

public static SttAssignResult Run(
    Database db, Editor ed, ObjectId polyId, string blockName,
    List<string> codes,
    List<string> drawingNames = null,
    bool allowMismatch = false)
{
    // ... PASS 1 giữ nguyên

    // PASS 2 — apply
    for (int i = 0; i < n; i++)
    {
        var br = (BlockReference)tr.GetObject(pending[i], OpenMode.ForWrite);
        FrameScanner.WriteStt(br, codes[i], cfg, tr);

        if (drawingNames != null && i < drawingNames.Count
            && !string.IsNullOrWhiteSpace(drawingNames[i]))
        {
            if (FrameScanner.WriteDrawingName(br, drawingNames[i], cfg, tr))
                result.DrawingNamesAssigned++;
            else
                result.DrawingNameAttributesMissing++;
        }
    }
    // ...
}
```

Giữ cả hai overload để code cũ không vỡ:

```csharp
public static SttAssignResult ApplyNumbersToSelectedFrames(
    Database db, Editor ed, ObjectId polyId, string blockName,
    List<string> codes, List<string> drawingNames, bool allowMismatch = false)
    => Run(db, ed, polyId, blockName, codes, drawingNames, allowMismatch);

public static SttAssignResult ApplyNumbersToSelectedFrames(
    Database db, Editor ed, ObjectId polyId, string blockName,
    List<string> codes, bool allowMismatch = false)
    => Run(db, ed, polyId, blockName, codes, null, allowMismatch);
```

### 2.4 `Core/NumberImporter.cs` — thêm cột DrawingName

```csharp
public class ImportedRow
{
    public int? Order;
    public string FrameNumber = "";
    public string DrawingName = "";
    public string Note = "";
    public int SourceLine;
}

public class ImportResult
{
    // ...
    public bool HasDrawingNameColumn;
}

private static readonly string[] DrawingNameAliases =
    { "drawingname", "drawing_name", "namedrawing", "name_drawing",
      "inno_name_drawing", "tênbảnvẽ", "tenbanve" };
```

Trong `ParseTabular`, dò cột và đọc giá trị:

```csharp
int colFrame = -1, colOrder = -1, colDrawingName = -1, colNote = -1;
// ...
else if (colDrawingName < 0 && MatchesAny(norm, DrawingNameAliases)) colDrawingName = c;
// ...
result.HasDrawingNameColumn = colDrawingName >= 0;

string drawingName = (colDrawingName >= 0 && colDrawingName < row.Length)
    ? (row[colDrawingName] ?? "").Trim()
    : "";
```

Template export đổi header thành `Order,FrameNumber,DrawingName,Note` (CSV) và
thêm ô `D1 = DrawingName` cho XLSX.

### 2.5 `UI/SttOptionsDialog.cs` — output, rule sinh tên, preview

```csharp
public List<string> Codes { get; private set; }
public List<string> DrawingNames { get; private set; }

private CheckBox _assignDrawingNameChk;
private TextBox _nPrefix, _nSuffix, _nSkip;
private NumericUpDown _nStart, _nStep, _nPadding;
```

Trong `DoPreview`, sau khi có `codes`:

```csharp
if (codes != null && _assignDrawingNameChk.Checked)
{
    if (mode == NumberingMode.Import && _importHasDrawingNames)
        drawingNames = new List<string>(_importedDrawingNames);
    else
        drawingNames = GenerateDrawingNames(codes.Count, errors);
}
```

Rule sinh tên dùng lại bộ sinh Simple:

```csharp
private List<string> GenerateDrawingNames(int count, List<string> errors)
{
    var p = new SimpleGenerationParams
    {
        Prefix  = _nPrefix.Text ?? "",
        Suffix  = _nSuffix.Text ?? "",
        Start   = (int)_nStart.Value,
        Step    = (int)_nStep.Value,
        Padding = (int)_nPadding.Value,
        Skip    = NumberGenerator.ParseSkipList(_nSkip.Text),
        Count   = count,
    };
    if (p.Step == 0) { errors.Add("Drawing name Step must not be zero."); return null; }
    return NumberGenerator.GenerateSimple(p);
}
```

Lưới preview thêm cột `DrawingNameCol`, và `Accept()` gán:

```csharp
DrawingNames = _cachedDrawingNames == null ? null : new List<string>(_cachedDrawingNames);
```

### 2.6 `Commands.cs` — PLSTT truyền tên, PLAUTO tạo đủ 2 attribute

```csharp
List<string> codes, drawingNames;
using (var dlg = new SttOptionsDialog(expectedCount))
{
    if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
    codes        = dlg.Codes;
    drawingNames = dlg.DrawingNames;
    allowMismatch = dlg.AllowCountMismatch;
}

var result = SttAssigner.ApplyNumbersToSelectedFrames(
    db, ed, polyId, choice.Name, codes, drawingNames, allowMismatch);

if (drawingNames != null)
{
    ed.WriteMessage(
        $"\nDrawing names assigned: {result.DrawingNamesAssigned}. " +
        $"Frames missing attribute '{Config.Instance.DrawingNameTag}': " +
        $"{result.DrawingNameAttributesMissing}.");
}
```

PLAUTO:

```csharp
NativeFrameBuilder.EnsureFrameBlock(
    db, nativeBlockName, w, h,
    Config.Instance.AttributeTag,
    Config.Instance.DrawingNameTag);
```

### 2.7 `Core/NativeFrameBuilder.EnsureFrameBlock` — tạo sẵn 2 attribute definition

Đổi chữ ký thêm tham số `drawingNameTag`, phần dựng attribute như sau:

```csharp
public static ObjectId EnsureFrameBlock(
    Database db,
    string blockName,
    double width,
    double height,
    string attributeTag,
    string drawingNameTag)
{
    // ... phần tạo/xoá BlockTableRecord và vẽ rectangle giữ nguyên

    // Attribute definitions — centred near each other, height ≈ 5% of frame height.
    double h = Math.Max(1.0, Math.Min(width, height) * 0.05);
    var att = new AttributeDefinition
    {
        Tag = attributeTag,
        Prompt = attributeTag,
        TextString = "",
        Position = new Point3d(width / 2.0, height / 2.0 + h, 0),
        Height = h,
        Justify = AttachmentPoint.MiddleCenter,
        AlignmentPoint = new Point3d(width / 2.0, height / 2.0 + h, 0),
        LockPositionInBlock = true,
    };
    btr.AppendEntity(att);
    tr.AddNewlyCreatedDBObject(att, true);

    if (!string.IsNullOrWhiteSpace(drawingNameTag)
        && !string.Equals(attributeTag, drawingNameTag, StringComparison.OrdinalIgnoreCase))
    {
        var nameAtt = new AttributeDefinition
        {
            Tag = drawingNameTag,
            Prompt = drawingNameTag,
            TextString = "",
            Position = new Point3d(width / 2.0, height / 2.0 - h, 0),
            Height = h,
            Justify = AttachmentPoint.MiddleCenter,
            AlignmentPoint = new Point3d(width / 2.0, height / 2.0 - h, 0),
            LockPositionInBlock = true,
        };
        btr.AppendEntity(nameAtt);
        tr.AddNewlyCreatedDBObject(nameAtt, true);
    }

    tr.Commit();
    return btrId;
}
```

`InsertFrames` không phải sửa: nó vốn duyệt mọi `AttributeDefinition` trong BTR nên
attribute tên tự động được tạo theo.

### 2.8 Bố cục dialog — các thay đổi về chỉ số hàng

`BuildLayout()` chèn thêm một hàng cho nhóm drawing name, nên `RowCount` tăng từ 5 lên 6
và mọi control phía dưới dịch xuống 1 hàng:

```csharp
RowCount = 6,
// ...
root.RowStyles.Add(new RowStyle(SizeType.Absolute, 420)); // tabs
root.RowStyles.Add(new RowStyle(SizeType.Absolute, 105)); // drawing names  <-- mới
root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));  // preview row
root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));  // validation
root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // grid
root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));  // buttons

root.Controls.Add(_tabs, 0, 0);
root.Controls.Add(BuildDrawingNamePanel(), 0, 1);
root.Controls.Add(previewPanel, 0, 2);
root.Controls.Add(_validation, 0, 3);
root.Controls.Add(_grid, 0, 4);
root.Controls.Add(btnPanel, 0, 5);
```

Form phóng to để đủ chỗ: `Height = 930`, `MinimumSize = new Size(800, 820)`.

Nhóm nhập rule tên bản vẽ (`BuildDrawingNamePanel`) là một `GroupBox` có tiêu đề hiển thị
luôn tag đang dùng, giúp người dùng biết đang ghi vào attribute nào:

```csharp
var group = new GroupBox
{
    Text = $"Drawing name attribute: {Config.Instance.DrawingNameTag}",
    Dock = DockStyle.Fill,
    Padding = new Padding(8),
};
```

Bên trong là `TableLayoutPanel` 8 cột × 2 hàng chứa checkbox `Assign drawing names`
(span 2 hàng) và các ô Prefix, Start, Step, Suffix, Pad, Skip.

Lưới preview đổi cột `Code` thành `Frame Number` (FillWeight 35), thêm cột
`DrawingNameCol` tiêu đề `Drawing Name` (FillWeight 40), cột `Note` giảm còn 25.

Khi import phát hiện cột DrawingName thì tự bật checkbox:

```csharp
if (import.HasDrawingNameColumn)
{
    sb.Append("  |  DrawingName column detected");
    _assignDrawingNameChk.Checked = true;
}
```

Các giá trị rule tên được lưu Registry cùng chỗ với cấu hình cũ
(`HKCU\Software\PrintLayoutAddin\Numbering`) bằng các khoá `AssignDrawingName`,
`NPrefix`, `NSuffix`, `NSkip`, `NStart`, `NStep`, `NPadding`.

---

## 3. Lưu ý vận hành

- Khung đã chèn sẵn trong bản vẽ: thêm attribute vào block gốc rồi chạy `ATTSYNC`,
  nếu không các instance cũ vẫn thiếu tag và sẽ bị đếm vào
  `DrawingNameAttributesMissing`.
- Đổi tên attribute cho khách: chỉ sửa `drawingNameTag` trong `config.json`
  (nhớ sửa cả bản trong `PrintLayoutAddin.bundle/Contents/net48` và `/net8`).
- Build: `dotnet build PrintLayoutAddin/PrintLayoutAddin.csproj -c Release`.
  Target net48 cần đường dẫn SDK AutoCAD, override bằng
  `/p:AutoCADPath="C:\Program Files\Autodesk\AutoCAD 2024"`.
