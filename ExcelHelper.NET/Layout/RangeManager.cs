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

    public RangeManager(ISheet sheet)
    {
        _sheet = sheet;
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
    /// Copy một vùng từ start cell đến end cell
    /// </summary>
    public void CopyRange(ICell startCell, ICell endCell, int targetStartRow)
    {
        var rowCount = endCell.RowIndex - startCell.RowIndex + 1;
        
        for (int i = 0; i < rowCount; i++)
        {
            _sheet.CloneRow(startCell.RowIndex + i, targetStartRow + i);
        }
    }

    /// <summary>
    /// Copy một vùng từ start cell đến end cell với xử lý merge regions
    /// </summary>
    public void CopyRangeWithMerge(ICell startCell, ICell endCell, int targetStartRow, List<CellRangeAddress> sourceMergeRegions)
    {
        var rowCount = endCell.RowIndex - startCell.RowIndex + 1;
        var sourceStartRow = startCell.RowIndex;
        
        // Copy từng row
        for (int i = 0; i < rowCount; i++)
        {
            _sheet.CloneRow(sourceStartRow + i, targetStartRow + i);
        }

        // Copy merge regions
        foreach (var sourceRegion in sourceMergeRegions)
        {
            var rowOffset = targetStartRow - sourceStartRow;
            var newRegion = new CellRangeAddress(
                sourceRegion.FirstRow + rowOffset,
                sourceRegion.LastRow + rowOffset,
                sourceRegion.FirstColumn,
                sourceRegion.LastColumn);

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
    /// Tạo nhiều bản copy của một range với xử lý merge cells (performance version)
    /// </summary>
    public void CreateRangesWithMerge(string sourceRange, int count)
    {
        var cells = sourceRange.Split(':');
        if (cells.Length != 2) return;

        var startCell = _sheet.GetCellByAddress(cells[0]);
        var endCell = _sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        var rangeRowCount = endCell.RowIndex - startCell.RowIndex + 1;
        var totalRowsNeeded = rangeRowCount * count;
        var rowsToMove = _sheet.LastRowNum - endCell.RowIndex;

        // Thu thập thông tin merge regions trong source range
        var sourceMergeRegions = new List<CellRangeAddress>();
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow >= startCell.RowIndex && region.LastRow <= endCell.RowIndex)
            {
                sourceMergeRegions.Add(region);
            }
        }

        // Di chuyển các dòng hiện có xuống dưới
        if (count > 1 && rowsToMove > 0)
        {
            for (int i = rowsToMove; i >= 1; i--)
            {
                _sheet.MoveRow(endCell.RowIndex + i, endCell.RowIndex + i + totalRowsNeeded - rangeRowCount);
            }
        }

        // Tạo các bản copy với merge regions
        var copyStartRow = endCell.RowIndex + 1;
        for (int i = 0; i < count - 1; i++) // Trừ 1 vì đã có bản gốc
        {
            var targetStartRow = copyStartRow + i * rangeRowCount;
            CopyRangeWithMerge(startCell, endCell, targetStartRow, sourceMergeRegions);
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
