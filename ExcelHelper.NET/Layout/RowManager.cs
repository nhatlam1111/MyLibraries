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
    /// Clone một dòng
    /// </summary>
    public void CloneRow(int sourceRowIndex, int targetRowIndex)
    {
        _sheet.CloneRow(sourceRowIndex, targetRowIndex);
    }


    /// <summary>
    /// Di chuyển một dòng
    /// </summary>
    public void MoveRow(int sourceRowIndex, int targetRowIndex)
    {
        _sheet.MoveRow(sourceRowIndex, targetRowIndex);
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
