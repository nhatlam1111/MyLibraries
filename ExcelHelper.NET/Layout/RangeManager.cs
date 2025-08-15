using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelHelper.NET.Extensions;
using ExcelHelper.NET.Utils;

namespace ExcelHelper.NET.Layout;

/// <summary>
/// Class quản lý các vùng cells và ranges
/// </summary>
public class RangeManager
{
    private readonly ISheet _sheet;
    private readonly MergeManager _mergeManager;

    public RangeManager(ISheet sheet)
    {
        _sheet = sheet;
        _mergeManager = new MergeManager(sheet);
    }

    /// <summary>
    /// Copy một vùng đến vị trí mới
    /// </summary>
    public void CopyRange(string sourceRange, int targetStartRow)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        CopyRange(startCell, endCell, targetStartRow);
    }

    /// <summary>
    /// Copy một vùng đến vị trí mới với chỉ định cột
    /// </summary>
    public void CopyRange(string sourceRange, int targetStartRow, int targetStartCol)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        CopyRange(startCell, endCell, targetStartRow, targetStartCol);
    }

    /// <summary>
    /// Copy một vùng từ start cell đến end cell
    /// </summary>
    public void CopyRange(ICell startCell, ICell endCell, int targetStartRow)
    {
        CopyRange(startCell, endCell, targetStartRow, startCell.ColumnIndex);
    }

    /// <summary>
    /// Copy một vùng từ start cell đến end cell với chỉ định vị trí đích
    /// </summary>
    public void CopyRange(ICell startCell, ICell endCell, int targetStartRow, int targetStartCol)
    {
        var rowCount = endCell.RowIndex - startCell.RowIndex + 1;
        var colCount = endCell.ColumnIndex - startCell.ColumnIndex + 1;
        
        for (int r = 0; r < rowCount; r++)
        {
            var sourceRow = _sheet.GetRow(startCell.RowIndex + r);
            if (sourceRow == null) continue;
            
            var targetRow = _sheet.GetRow(targetStartRow + r) ?? _sheet.CreateRow(targetStartRow + r);
            
            for (int c = 0; c < colCount; c++)
            {
                var sourceCell = sourceRow.GetCell(startCell.ColumnIndex + c);
                var targetCell = targetRow.GetCell(targetStartCol + c) ?? targetRow.CreateCell(targetStartCol + c);
                
                if (sourceCell != null)
                {
                    // Sử dụng helper method để copy giá trị cell
                    _sheet.CopyCellValue(sourceCell, targetCell);
                    targetCell.CellStyle = sourceCell.CellStyle;
                }
            }
        }
    }

    /// <summary>
    /// Copy một vùng từ start cell đến end cell với xử lý merge regions
    /// </summary>
    public void CopyRangeWithMerge(ICell startCell, ICell endCell, int targetStartRow)
    {
        CopyRangeWithMerge(startCell, endCell, targetStartRow, startCell.ColumnIndex);
    }

    /// <summary>
    /// Copy một vùng từ start cell đến end cell với xử lý merge regions và chỉ định cột đích
    /// </summary>
    public void CopyRangeWithMerge(ICell startCell, ICell endCell, int targetStartRow, int targetStartCol)
    {
        var sourceStartRow = startCell.RowIndex;
        var sourceEndRow = endCell.RowIndex;
        var sourceStartCol = startCell.ColumnIndex;
        var sourceEndCol = endCell.ColumnIndex;
        
        // Copy cells với vị trí mới
        CopyRange(startCell, endCell, targetStartRow, targetStartCol);

        // Lấy merge regions trong source range và copy với offset
        var sourceMergeRegions = _mergeManager.GetMergeRegionsInRange(
            sourceStartRow, sourceEndRow, sourceStartCol, sourceEndCol);

        foreach (var sourceRegion in sourceMergeRegions)
        {
            var rowOffset = targetStartRow - sourceStartRow;
            var colOffset = targetStartCol - sourceStartCol;
            
            var newRegion = new CellRangeAddress(
                sourceRegion.FirstRow + rowOffset,
                sourceRegion.LastRow + rowOffset,
                sourceRegion.FirstColumn + colOffset,
                sourceRegion.LastColumn + colOffset);

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

    /// <summary>
    /// Copy một vùng đến vị trí mới với xử lý merge regions
    /// </summary>
    public void CopyRangeWithMerge(string sourceRange, int targetStartRow)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        CopyRangeWithMerge(startCell, endCell, targetStartRow);
    }

    /// <summary>
    /// Copy một vùng đến vị trí mới với xử lý merge regions và chỉ định cột
    /// </summary>
    public void CopyRangeWithMerge(string sourceRange, int targetStartRow, int targetStartCol)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        CopyRangeWithMerge(startCell, endCell, targetStartRow, targetStartCol);
    }

    /// <summary>
    /// Tạo nhiều bản copy của một range
    /// </summary>
    public void CreateRanges(string sourceRange, int count)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        var rangeRowCount = endCell.RowIndex - startCell.RowIndex + 1;
        var totalRowsNeeded = rangeRowCount * count;
        var rowsToMove = _sheet.LastRowNum - endCell.RowIndex;

        // Di chuyển các dòng hiện có xuống dưới
        if (count > 1 && rowsToMove > 0)
        {
            for (int i = rowsToMove; i >= 1; i--)
            {
                _sheet.MoveRow(endCell.RowIndex + i, endCell.RowIndex + i + totalRowsNeeded - rangeRowCount);
            }
        }

        // Tạo các bản copy
        var copyStartRow = endCell.RowIndex + 1;
        for (int i = 0; i < count - 1; i++) // Trừ 1 vì đã có bản gốc
        {
            CopyRange(startCell, endCell, copyStartRow + i * rangeRowCount);
        }
    }

    /// <summary>
    /// Clear nội dung của một range
    /// </summary>
    public void ClearRange(string range)
    {
        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        for (int row = startCell.RowIndex; row <= endCell.RowIndex; row++)
        {
            for (int col = startCell.ColumnIndex; col <= endCell.ColumnIndex; col++)
            {
                var cell = _sheet.GetCellByIndex(row, col);
                cell.SetBlank();
            }
        }
    }

    /// <summary>
    /// Đặt giá trị cho tất cả cells trong range
    /// </summary>
    public void SetRangeValue(string range, object value)
    {
        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        for (int row = startCell.RowIndex; row <= endCell.RowIndex; row++)
        {
            for (int col = startCell.ColumnIndex; col <= endCell.ColumnIndex; col++)
            {
                var cell = _sheet.GetCellByIndex(row, col);
                cell.SetValue(value);
            }
        }
    }

    /// <summary>
    /// Lấy tất cả cells trong một range
    /// </summary>
    public List<ICell> GetCellsInRange(string range)
    {
        var result = new List<ICell>();
        var cells = range.Split(':');
        if (cells.Length != 2) return result;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return result;

        for (int row = startCell.RowIndex; row <= endCell.RowIndex; row++)
        {
            for (int col = startCell.ColumnIndex; col <= endCell.ColumnIndex; col++)
            {
                var cell = _sheet.GetCellByIndex(row, col);
                result.Add(cell);
            }
        }

        return result;
    }
}
