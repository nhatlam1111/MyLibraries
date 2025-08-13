using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
                             InsertOptions? options = null)
    {
        var dataList = data.ToList();
        if (!dataList.Any()) return;

        dataProvider ??= new GenericDataProvider<T>();
        options ??= new InsertOptions();

        var rowsNeeded = dataList.Count;
        
        // Tạo đủ số dòng cần thiết
        _rowManager.CreateRows(startRow, rowsNeeded);

        // Lấy thông tin các cột từ header row
        var headerRow = _sheet.GetRow(startRow);
        if (headerRow == null) return;

        var columnMappings = GetColumnMappings(headerRow, dataProvider);

        // Điền dữ liệu
        var sequenceNumber = options.SequenceStartNumber;
        for (int i = 0; i < dataList.Count; i++)
        {
            var rowIndex = startRow + i;
            FillRowData(dataList[i], rowIndex, columnMappings, dataProvider, 
                       options.InsertSequence ? sequenceNumber++ : null);

            if (options.AutofitHeight)
            {
                _rowManager.AutoFitRowHeight(rowIndex);
            }
        }
    }

    /// <summary>
    /// Chèn dữ liệu từ DataTable (sử dụng adapter)
    /// </summary>
    public void InsertData<T>(DataTable dataTable, int startRow, InsertOptions? options = null) 
        where T : new()
    {
        var data = DataTableAdapter.ToList<T>(dataTable);
        var dataProvider = new GenericDataProvider<T>();
        InsertData(data, startRow, dataProvider, options);
    }

    /// <summary>
    /// Chèn dữ liệu từ DataTable vào range với merge support
    /// </summary>
    public void InsertDataToRange<T>(DataTable dataTable, string range, bool preserveMerge = true) 
        where T : new()
    {
        if (dataTable == null || dataTable.Rows.Count == 0) return;

        // Convert DataTable to List<T> và sử dụng hàm InsertDataToRange<T> đã có
        var data = DataTableAdapter.ToList<T>(dataTable);
        var dataProvider = new GenericDataProvider<T>();
        
        InsertDataToRange(data, range, dataProvider, preserveMerge);
    }

    /// <summary>
    /// Chèn dữ liệu vào một range cụ thể
    /// </summary>
    public void InsertDataToRange<T>(IEnumerable<T> data, string range,
                                    IExcelDataProvider<T>? dataProvider = null, bool preserveMerge = true)
    {
        var dataList = data.ToList();
        if (!dataList.Any()) return;

        dataProvider ??= new GenericDataProvider<T>();

        // Tạo các range copies với hoặc không merge
        if (preserveMerge)
        {
            _rangeManager.CreateRangesWithMerge(range, dataList.Count);
        }
        else
        {
            _rangeManager.CreateRanges(range, dataList.Count);
        }

        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        var rangeRowCount = endCell.RowIndex - startCell.RowIndex + 1;

        // Fill data vào từng range
        for (int i = 0; i < dataList.Count; i++)
        {
            var rangeStartRow = startCell.RowIndex + i * rangeRowCount;
            FillRangeData(dataList[i], rangeStartRow, rangeStartRow + rangeRowCount - 1, dataProvider);
        }
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
            
            // Tìm field name match với cell value
            var matchingField = fieldNames.FirstOrDefault(f => 
                f.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
                
            if (!string.IsNullOrEmpty(matchingField))
            {
                mappings[i] = matchingField;
            }
        }

        return mappings;
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

            // Xử lý sequence number
            if (sequenceNumber.HasValue && 
                (fieldName.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                 fieldName.Equals("stt", StringComparison.OrdinalIgnoreCase)))
            {
                cell.SetValue(sequenceNumber.Value);
                continue;
            }

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

    /// <summary>
    /// Copy một range đến vị trí khác với xử lý merge cells
    /// </summary>
    public void CopyRangeWithMerge(string sourceRange, int targetStartRow)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        // Thu thập merge regions trong source range
        var sourceMergeRegions = _mergeManager.GetMergeRegionsInRange(
            startCell.RowIndex, endCell.RowIndex, 
            startCell.ColumnIndex, endCell.ColumnIndex);

        // Copy range với merge
        _rangeManager.CopyRangeWithMerge(startCell, endCell, targetStartRow, sourceMergeRegions);
    }

    /// <summary>
    /// Tạo nhiều rows từ template với hỗ trợ merge cells
    /// </summary>
    public void CreateRowsWithMerge(int templateRowIndex, int count, bool moveExistingRows = true)
    {
        _rowManager.CreateRowsWithMerge(templateRowIndex, count, moveExistingRows);
    }

    /// <summary>
    /// Move row với hỗ trợ merge cells
    /// </summary>
    public void MoveRowWithMerge(int sourceRowIndex, int targetRowIndex)
    {
        _rowManager.MoveRowWithMerge(sourceRowIndex, targetRowIndex);
    }
}
