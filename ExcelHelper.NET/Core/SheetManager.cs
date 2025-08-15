using System.Data;
using System.Text.RegularExpressions;
using ExcelHelper.NET.Data;
using ExcelHelper.NET.Extensions;
using ExcelHelper.NET.Layout;
using ExcelHelper.NET.Media;
using ExcelHelper.NET.Models;
using ExcelHelper.NET.Styling;

namespace ExcelHelper.NET.Core;

/// <summary>
/// Class quản lý các thao tác trên sheet
/// </summary>
public class SheetManager
{
    private readonly ISheet _sheet;
    private readonly XSSFWorkbook _workbook;
    private readonly RowManager _rowManager;
    private readonly RangeManager _rangeManager;
    private readonly MergeManager _mergeManager;
    private readonly ImageManager _imageManager;
    private readonly CellStyler _cellStyler;

    public SheetManager(ISheet sheet)
    {
        _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        _workbook = (XSSFWorkbook)sheet.Workbook;
        
        _rowManager = new RowManager(_sheet);
        _rangeManager = new RangeManager(_sheet);
        _mergeManager = new MergeManager(_sheet);
        _imageManager = new ImageManager(_sheet);
        _cellStyler = new CellStyler(_workbook);
    }

    /// <summary>
    /// Row Manager
    /// </summary>
    public RowManager Rows => _rowManager;

    /// <summary>
    /// Range Manager
    /// </summary>
    public RangeManager Ranges => _rangeManager;

    /// <summary>
    /// Merge Manager
    /// </summary>
    public MergeManager Merges => _mergeManager;

    /// <summary>
    /// Image Manager
    /// </summary>
    public ImageManager Images => _imageManager;

    /// <summary>
    /// Cell Styler
    /// </summary>
    public CellStyler Styles => _cellStyler;

    /// <summary>
    /// Sheet hiện tại
    /// </summary>
    public ISheet Sheet => _sheet;

    /// <summary>
    /// Chèn dữ liệu generic từ List&lt;T&gt;
    /// </summary>
    public void InsertData<T>(IEnumerable<T> data, int startRow, 
                             IExcelDataProvider<T>? dataProvider = null,
                             InsertOptions? options = null, bool moveExistingRows = true)
    {
        var dataList = data.ToList();
        if (!dataList.Any()) return;

        var rowsNeeded = dataList.Count;
        dataProvider ??= new GenericDataProvider<T>();
        options ??= new InsertOptions();
        
        // Lấy thông tin các cột từ header row
        var headerRow = _sheet.GetRow(startRow);
        if (headerRow == null) return;

        var columnMappings = GetColumnMappings(headerRow, dataProvider);

        var templateMergeRegions = new List<MergeRegion>();
        var beforeTemplateMergeRegions = new List<MergeRegion>();
        var afterTemplateMergeRegions = new List<MergeRegion>();

        //lưu lai merge list hiện tại xử lý sau và xóa toàn bộ merge regions hiện tại để tối ưu hóa tốc độ
        List<MergeRegion> mergeRegions = _mergeManager.GetAllMergeRegions();
        _mergeManager.ClearAllMergeRegions();

        // Tìm merge regions
        foreach (var mr in mergeRegions)
        {
            if (mr.FirstRow == startRow) templateMergeRegions.Add(mr);
            else if (mr.FirstRow < startRow) beforeTemplateMergeRegions.Add(mr);
            else afterTemplateMergeRegions.Add(mr);
        }

        // Tạo đủ số dòng cần thiết
        _rowManager.CreateRows(startRow, rowsNeeded, moveExistingRows);

        // Điền dữ liệu
        var sequenceNumber = options.SequenceStartNumber;
        for (int i = 0; i < dataList.Count; i++)
        {
            var rowIndex = startRow + i;
            FillRowData(dataList[i], rowIndex, columnMappings, dataProvider, 
                       options.InsertSequence ? sequenceNumber++ : null);

            // Khôi phục merge regions cho dòng hiện tại nếu template có merge
            if (templateMergeRegions.Any()) 
            {
                _mergeManager.RestoreMergeRegions(templateMergeRegions, i);
            }

            if (options.AutofitHeight)
            {
                _rowManager.AutoFitRowHeight(rowIndex);
            }
        }

        // Khôi phục merge regions đoạn trước template và sau template
        _mergeManager.RestoreMergeRegions(beforeTemplateMergeRegions, 0); // Không shift
        _mergeManager.RestoreMergeRegions(afterTemplateMergeRegions, rowsNeeded - 1); // Shift theo số rows inserted
    }

    /// <summary>
    /// Chèn dữ liệu từ DataTable (sử dụng adapter)
    /// </summary>
    public void InsertData<T>(DataTable dataTable, int startRow, InsertOptions? options = null, bool moveExistingRows = true) 
        where T : new()
    {
        var data = DataTableAdapter.ToList<T>(dataTable);
        var dataProvider = new GenericDataProvider<T>();
        InsertData(data, startRow, dataProvider, options, moveExistingRows);
    }

    /// <summary>
    /// Chèn dữ liệu từ DataTable vào range với merge support
    /// </summary>
    public void InsertDataToRange<T>(DataTable dataTable, string range) 
        where T : new()
    {
        if (dataTable == null || dataTable.Rows.Count == 0) return;

        // Convert DataTable to List<T> và sử dụng hàm InsertDataToRange<T> đã có
        var data = DataTableAdapter.ToList<T>(dataTable);
        var dataProvider = new GenericDataProvider<T>();
        
        InsertDataToRange(data, range, dataProvider);
    }

    /// <summary>
    /// Chèn dữ liệu vào một range cụ thể
    /// </summary>
    public void InsertDataToRange<T>(IEnumerable<T> data, string range,
                                    IExcelDataProvider<T>? dataProvider = null)
    {
        var dataList = data.ToList();
        if (!dataList.Any()) return;

        dataProvider ??= new GenericDataProvider<T>();

        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        var rangeRowCount = endCell.RowIndex - startCell.RowIndex + 1;

        // Lưu và clear merge regions để tối ưu tốc độ (tương tự InsertData)
        List<MergeRegion> mergeRegions = new List<MergeRegion>();
        var templateMergeRegions = new List<MergeRegion>();
        var otherMergeRegions = new List<MergeRegion>();
        mergeRegions = _mergeManager.GetAllMergeRegions();
        _mergeManager.ClearAllMergeRegions();

        foreach (var mr in mergeRegions)
        {
            if (mr.FirstRow >= startCell.RowIndex && mr.LastRow <= endCell.RowIndex &&
                        mr.FirstColumn >= startCell.ColumnIndex && mr.LastColumn <= endCell.ColumnIndex)
            {
                templateMergeRegions.Add(mr);
            }
            else
            {
                otherMergeRegions.Add(mr);
            }
        }
       
        // Tạo các range copies
        _rangeManager.CreateRanges(range, dataList.Count);

        // Fill data vào từng range
        for (int i = 0; i < dataList.Count; i++)
        {
            var rangeStartRow = startCell.RowIndex + i * rangeRowCount;
            FillRangeData(dataList[i], rangeStartRow, rangeStartRow + rangeRowCount - 1, dataProvider);

            // Copy merge regions cho range thứ i (trừ range đầu tiên vì đã có sẵn)
            if (templateMergeRegions.Any())
            {
                var rowOffset = i * rangeRowCount;
                _mergeManager.RestoreMergeRegions(templateMergeRegions, rowOffset);
            }
        }
        
        // Khôi phục merge regions khác
        _mergeManager.RestoreMergeRegions(otherMergeRegions, 0);
    }

    /// <summary>
    /// Lấy mapping giữa column và field names
    /// </summary>
    private Dictionary<int, string> GetColumnMappings<T>(IRow headerRow, IExcelDataProvider<T> dataProvider)
    {
        var mappings = new Dictionary<int, string>();
        var fieldNames = dataProvider.GetFieldNames().ToList();

        for (int i = 0; i < headerRow.LastCellNum; i++)
        {
            var cell = headerRow.GetCell(i);
            if (cell == null) continue;

            var cellValue = cell.GetStringValue().Trim().ToLower();
            
            // Ưu tiên map field từ T trước
            var matchingField = fieldNames.FirstOrDefault(f => 
                f.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
                
            if (!string.IsNullOrEmpty(matchingField))
            {
                mappings[i] = matchingField;
            }
            // Nếu không có trong T, check xem có phải sequence field không
            else if (IsSequenceField(cellValue))
            {
                mappings[i] = cellValue; // Map trực tiếp để xử lý sequence
            }
        }

        return mappings;
    }

    /// <summary>
    /// Kiểm tra field có phải sequence field (STT, NO) không
    /// </summary>
    private static bool IsSequenceField(string fieldName)
    {
        return fieldName.Equals("no", StringComparison.OrdinalIgnoreCase) ||
               fieldName.Equals("stt", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Fill dữ liệu vào một dòng
    /// </summary>
    private void FillRowData<T>(T item, int rowIndex, Dictionary<int, string> columnMappings,
                               IExcelDataProvider<T> dataProvider, int? sequenceNumber)
    {
        var row = _sheet.GetRow(rowIndex);
        if (row == null) return;

        foreach (var (columnIndex, fieldName) in columnMappings)
        {
            var cell = row.GetCell(columnIndex);
            if (cell == null) continue;

            // Xử lý sequence number (ưu tiên sequence trước khi lấy field value từ T)
            if (sequenceNumber.HasValue && IsSequenceField(fieldName))
            {
                cell.SetValue(sequenceNumber.Value);
                continue;
            }

            // Nếu không phải sequence field, lấy field value từ T
            var fieldValue = dataProvider.GetFieldValue(item, fieldName);
            
            // Xử lý hình ảnh
            if (dataProvider.IsImageField(fieldName) && fieldValue is byte[] imageData)
            {
                _imageManager.InsertImage(cell, imageData);
                continue;
            }

            // Xử lý các kiểu dữ liệu khác
            cell.SetValue(fieldValue);
        }
    }

    /// <summary>
    /// Fill dữ liệu vào một range với template variables
    /// </summary>
    private void FillRangeData<T>(T item, int rangeStartRow, int rangeEndRow, IExcelDataProvider<T> dataProvider)
    {
        for (int rowIndex = rangeStartRow; rowIndex <= rangeEndRow; rowIndex++)
        {
            var row = _sheet.GetRow(rowIndex);
            if (row == null) continue;

            for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
            {
                var cell = row.GetCell(colIndex);
                if (cell == null || cell.CellType != CellType.String) continue;

                var cellText = cell.StringCellValue;
                if (string.IsNullOrEmpty(cellText)) continue;

                // Tìm và thay thế các template variables có dạng $[FieldName]
                var updatedText = ReplaceTemplateVariables(cellText, item, dataProvider);
                if (updatedText != cellText)
                {
                    cell.SetCellValue(updatedText);
                }
            }
        }
    }

    /// <summary>
    /// Thay thế template variables trong text
    /// </summary>
    private string ReplaceTemplateVariables<T>(string text, T item, IExcelDataProvider<T> dataProvider)
    {
        var regex = new Regex(@"\$\[(.*?)\]");
        var matches = regex.Matches(text);
        
        var result = text;
        foreach (Match match in matches)
        {
            var fieldName = match.Groups[1].Value.Trim();
            var fieldValue = dataProvider.GetFieldValue(item, fieldName);
            var stringValue = fieldValue?.ToString() ?? "";
            
            result = result.Replace(match.Value, stringValue);
        }
        
        return result;
    }

    /// <summary>
    /// Áp dụng CellFormat cho một cell
    /// </summary>
    public void ApplyCellFormat(int rowIndex, int columnIndex, CellFormat format)
    {
        var cell = _sheet.GetCellByIndex(rowIndex, columnIndex);
        var style = _cellStyler.CreateStyle(format);
        cell.CellStyle = style;
    }

    /// <summary>
    /// Áp dụng CellFormat cho một range
    /// </summary>
    public void ApplyCellFormat(string range, CellFormat format)
    {
        var cells = _rangeManager.GetCellsInRange(range);
        var style = _cellStyler.CreateStyle(format);
        
        foreach (var cell in cells)
        {
            cell.CellStyle = style;
        }
    }

    /// <summary>
    /// Xóa một cột
    /// </summary>
    public void DeleteColumn(int columnIndex)
    {
        _sheet.DeleteColumn(columnIndex);
    }
}
