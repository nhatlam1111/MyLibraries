# ExcelHelper.NET - Merge Cell Handling Features

## Tổng quan

Dự án ExcelHelper.NET đã được bổ sung các tính năng xử lý merge cells khi copy range, move row và tạo ranges. Điều này khắc phục vấn đề từ version cũ trong file ExcelHelper.txt.

## Các tính năng mới được thêm

### 1. RangeManager - Xử lý Range với Merge Cells

#### `CreateRangesWithMerge(string sourceRange, int count)`
- Tạo nhiều bản copy của một range với bảo toàn merge cells
- Thu thập thông tin merge regions trong source range
- Copy merge regions với offset phù hợp cho từng target range

#### `CopyRangeWithMerge(ICell startCell, ICell endCell, int targetStartRow, List<CellRangeAddress> sourceMergeRegions)`
- Copy một range với danh sách merge regions được cung cấp
- Tính toán offset và tạo merge regions mới tại vị trí target

### 2. RowManager - Xử lý Row với Merge Cells

#### `CreateRowsWithMerge(int templateRowIndex, int count, bool moveExistingRows = true)`
- Tạo nhiều rows từ template với bảo toàn merge cells
- Thu thập merge regions của template row
- Áp dụng merge cho tất cả rows được tạo

#### `CloneRowWithMerge(int sourceRowIndex, int targetRowIndex)`
- Clone một row với xử lý merge cells
- Tìm các merge regions chứa source row
- Tạo merge regions tương ứng cho target row

#### `MoveRowWithMerge(int sourceRowIndex, int targetRowIndex)`
- Di chuyển row với cập nhật merge regions
- Thu thập các merge regions bị ảnh hưởng
- Cập nhật vị trí merge regions sau khi move

### 3. MergeManager - Quản lý Merge Regions

#### `GetMergeRegionsInRange(int startRow, int endRow, int startCol, int endCol)`
- Lấy tất cả merge regions trong một vùng cụ thể
- Hỗ trợ cho các operations copy range

#### `CopyMergeRegions(int sourceStartRow, int sourceEndRow, int sourceStartCol, int sourceEndCol, int targetStartRow, int targetStartCol)`
- Copy merge regions từ source range đến target range
- Tính toán offset row và column

#### `ShiftMergeRegions(int startRow, int rowShift)`
- Dịch chuyển merge regions khi insert/delete rows
- Cập nhật vị trí của tất cả regions sau startRow

### 4. SheetManager - Interface cấp cao

#### `InsertDataToRange<T>(IEnumerable<T> data, string range, IExcelDataProvider<T>? dataProvider = null, bool preserveMerge = true)`
- Chèn dữ liệu vào range với tùy chọn bảo toàn merge cells
- Parameter `preserveMerge` cho phép bật/tắt xử lý merge

#### `InsertDataToRange<T>(DataTable dataTable, string range, bool preserveMerge = true) where T : new()`
- Chèn DataTable vào range với merge support
- Convert DataTable thành `List<T>` và tái sử dụng logic của `InsertDataToRange<T>`
- Đơn giản hóa code và tránh trùng lặp logic

#### `CopyRangeWithMerge(string sourceRange, int targetStartRow)`
- Copy range với merge cells ở mức SheetManager
- Tự động thu thập merge regions và copy

#### `CreateRowsWithMerge(int templateRowIndex, int count, bool moveExistingRows = true)`
- Wrapper method cho RowManager.CreateRowsWithMerge

#### `MoveRowWithMerge(int sourceRowIndex, int targetRowIndex)`
- Wrapper method cho RowManager.MoveRowWithMerge

## Template Variables

Hệ thống hỗ trợ template variables với format `$[FieldName]` hoặc `$[ColumnName]`:

- Cho generic data: `$[PropertyName]` sẽ được thay bằng giá trị property
- Cho DataTable: `$[ColumnName]` sẽ được thay bằng giá trị column

## Ví dụ sử dụng

```csharp
// Tạo SheetManager
var sheetManager = new SheetManager(sheet);

// Case 1: Insert data to range với merge preservation
var testData = new List<Employee> { /* data */ };
sheetManager.InsertDataToRange(testData, "A1:D5", preserveMerge: true);

// Case 2: Copy range với merge cells
sheetManager.CopyRangeWithMerge("A1:E10", targetStartRow: 15);

// Case 3: Create rows từ template với merge
sheetManager.CreateRowsWithMerge(templateRowIndex: 0, count: 5);

// Case 4: Move row với merge handling
sheetManager.MoveRowWithMerge(sourceRowIndex: 2, targetRowIndex: 6);

// Case 5: Insert DataTable với merge - sử dụng generic type
DataTable dataTable = CreateSampleDataTable();
sheetManager.InsertDataToRange<SampleData>(dataTable, "A1:D8", preserveMerge: true);
```

## So sánh với version cũ

### Version cũ (ExcelHelper.txt)
- Method `CreateRanges2()` với xử lý merge cứng nhắc
- Method `CloneRow2()` và `CloneRow3()` phức tạp
- Xử lý merge regions được hardcode trong từng method

### Version mới (ExcelHelper.NET)
- Architecture phân tách rõ ràng: RangeManager, RowManager, MergeManager
- API đơn giản và flexible
- Support generic data types và DataTable
- Tùy chọn bật/tắt merge handling
- Template variables system
- **Code reuse**: DataTable method tái sử dụng logic của generic method

## Performance

- Sử dụng `CreateRangesWithMerge` thay vì `CreateRanges` khi cần preserve merge
- Method thu thập merge regions một lần và reuse cho multiple copies
- Tối ưu thứ tự xóa/thêm merge regions để tránh index conflicts

## Error Handling

- Try-catch cho AddMergedRegion để tránh duplicate regions
- Validate range format và cell addresses
- Safe handling khi merge regions invalid hoặc out of bounds
