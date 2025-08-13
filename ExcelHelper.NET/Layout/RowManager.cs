using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelHelper.NET.Extensions;

namespace ExcelHelper.NET.Layout;

/// <summary>
/// Class quản lý các thao tác với dòng
/// </summary>
public class RowManager
{
    private readonly ISheet _sheet;

    public RowManager(ISheet sheet)
    {
        _sheet = sheet;
    }

    /// <summary>
    /// Tạo nhiều dòng từ template
    /// </summary>
    public void CreateRows(int templateRowIndex, int count, bool moveExistingRows = true)
    {
        if (count <= 1) return;
        
        _sheet.CreateRowsFromTemplate(templateRowIndex, count, moveExistingRows);
    }

    /// <summary>
    /// Tạo nhiều dòng từ template với xử lý merge cells
    /// </summary>
    public void CreateRowsWithMerge(int templateRowIndex, int count, bool moveExistingRows = true)
    {
        if (count <= 1) return;

        var rowsToInsert = count - 1; // Trừ 1 vì đã có template row
        var rowsToMove = _sheet.LastRowNum - templateRowIndex;

        // Thu thập merge regions của template row
        var templateMergeRegions = new List<CellRangeAddress>();
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow <= templateRowIndex && templateRowIndex <= region.LastRow)
            {
                templateMergeRegions.Add(region);
            }
        }

        // Di chuyển các dòng hiện có xuống dưới nếu cần
        if (moveExistingRows && rowsToMove > 0)
        {
            for (int i = rowsToMove; i >= 1; i--)
            {
                _sheet.MoveRow(templateRowIndex + i, templateRowIndex + i + rowsToInsert);
            }
        }

        // Tạo các dòng mới từ template
        for (int i = 0; i < rowsToInsert; i++)
        {
            var newRowIndex = templateRowIndex + i + 1;
            CloneRowWithMerge(templateRowIndex, newRowIndex);
        }
    }

    /// <summary>
    /// Clone một dòng
    /// </summary>
    public void CloneRow(int sourceRowIndex, int targetRowIndex)
    {
        _sheet.CloneRow(sourceRowIndex, targetRowIndex);
    }

    /// <summary>
    /// Clone một dòng với xử lý merge cells
    /// </summary>
    public void CloneRowWithMerge(int sourceRowIndex, int targetRowIndex)
    {
        // Clone row data
        _sheet.CloneRow(sourceRowIndex, targetRowIndex);

        // Handle merge regions for this row
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            
            // Kiểm tra nếu merge region chứa source row
            if (region.FirstRow <= sourceRowIndex && sourceRowIndex <= region.LastRow)
            {
                // Tính toán offset cho target row
                var rowOffset = targetRowIndex - sourceRowIndex;
                
                // Tạo merge region mới cho target row
                var newRegion = new CellRangeAddress(
                    region.FirstRow + rowOffset,
                    region.LastRow + rowOffset,
                    region.FirstColumn,
                    region.LastColumn);

                try
                {
                    _sheet.AddMergedRegion(newRegion);
                }
                catch (ArgumentException)
                {
                    // Region already exists or invalid, ignore
                }
            }
        }
    }

    /// <summary>
    /// Di chuyển một dòng
    /// </summary>
    public void MoveRow(int sourceRowIndex, int targetRowIndex)
    {
        _sheet.MoveRow(sourceRowIndex, targetRowIndex);
    }

    /// <summary>
    /// Di chuyển một dòng với xử lý merge cells
    /// </summary>
    public void MoveRowWithMerge(int sourceRowIndex, int targetRowIndex)
    {
        // Thu thập merge regions liên quan đến source row
        var affectedRegions = new List<(int index, CellRangeAddress region)>();
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow <= sourceRowIndex && sourceRowIndex <= region.LastRow)
            {
                affectedRegions.Add((i, region));
            }
        }

        // Di chuyển row
        _sheet.MoveRow(sourceRowIndex, targetRowIndex);

        // Cập nhật merge regions
        foreach (var (index, region) in affectedRegions.OrderByDescending(x => x.index))
        {
            // Xóa region cũ
            _sheet.RemoveMergedRegion(index);

            // Tính toán vị trí mới
            var rowOffset = targetRowIndex - sourceRowIndex;
            var newRegion = new CellRangeAddress(
                region.FirstRow + rowOffset,
                region.LastRow + rowOffset,
                region.FirstColumn,
                region.LastColumn);

            // Thêm region mới
            try
            {
                _sheet.AddMergedRegion(newRegion);
            }
            catch (ArgumentException)
            {
                // Region invalid, ignore
            }
        }
    }

    /// <summary>
    /// Xóa một dòng
    /// </summary>
    public void DeleteRow(int rowIndex)
    {
        var row = _sheet.GetRow(rowIndex);
        if (row != null)
        {
            _sheet.RemoveRow(row);
        }
    }

    /// <summary>
    /// Ẩn/hiện một dòng
    /// </summary>
    public void SetRowVisibility(int rowIndex, bool isHidden)
    {
        var row = _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);
        row.Hidden = isHidden;
    }

    /// <summary>
    /// Đặt chiều cao cho dòng
    /// </summary>
    public void SetRowHeight(int rowIndex, float heightInPoints)
    {
        var row = _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);
        row.HeightInPoints = heightInPoints;
    }

    /// <summary>
    /// Autofit chiều cao của dòng
    /// </summary>
    public void AutoFitRowHeight(int rowIndex)
    {
        var row = _sheet.GetRow(rowIndex);
        if (row != null)
        {
            // Calculate height based on content
            float maxHeight = row.HeightInPoints;
            
            for (int i = 0; i < row.LastCellNum; i++)
            {
                var cell = row.GetCell(i);
                if (cell != null && !cell.IsEmpty())
                {
                    var cellValue = cell.GetStringValue();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        // Simple height calculation based on line breaks
                        var lineCount = cellValue.Split('\n').Length;
                        var calculatedHeight = lineCount * 15; // 15 points per line
                        maxHeight = Math.Max(maxHeight, calculatedHeight);
                    }
                }
            }
            
            row.HeightInPoints = maxHeight;
        }
    }
}
