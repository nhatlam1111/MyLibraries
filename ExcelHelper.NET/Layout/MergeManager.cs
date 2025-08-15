using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelHelper.NET.Models;

namespace ExcelHelper.NET.Layout;

/// <summary>
/// Class quản lý các vùng merge cells
/// </summary>
public class MergeManager
{
    private readonly ISheet _sheet;

    public MergeManager(ISheet sheet)
    {
        _sheet = sheet;
    }

    /// <summary>
    /// Merge các cells trong một range
    /// </summary>
    public void MergeRange(string range)
    {
        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var (startRow, startCol) = Utils.ExcelAddressConverter.AddressToIndices(cells[0]);
        var (endRow, endCol) = Utils.ExcelAddressConverter.AddressToIndices(cells[1]);
    }

    /// <summary>
    /// Merge cells từ tọa độ cụ thể
    /// </summary>
    public void MergeCells(int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        try
        {
            var region = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
            _sheet.AddMergedRegion(region);
        }
        catch (ArgumentException)
        {
            // Region already merged or invalid, ignore
        }
    }

    /// <summary>
    /// Unmerge một vùng cells
    /// </summary>
    public void UnmergeRange(string range)
    {
        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var (startRow, startCol) = Utils.ExcelAddressConverter.AddressToIndices(cells[0]);
        var (endRow, endCol) = Utils.ExcelAddressConverter.AddressToIndices(cells[1]);
        
        UnmergeCells(startRow, endRow, startCol, endCol);
    }

    /// <summary>
    /// Unmerge cells từ tọa độ cụ thể
    /// </summary>
    public void UnmergeCells(int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        for (int i = _sheet.NumMergedRegions - 1; i >= 0; i--)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow == firstRow && region.LastRow == lastRow &&
                region.FirstColumn == firstColumn && region.LastColumn == lastColumn)
            {
                _sheet.RemoveMergedRegion(i);
                break;
            }
        }
    }

    /// <summary>
    /// Kiểm tra một cell có nằm trong vùng merge không
    /// </summary>
    public bool IsCellMerged(int rowIndex, int columnIndex)
    {
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow <= rowIndex && rowIndex <= region.LastRow &&
                region.FirstColumn <= columnIndex && columnIndex <= region.LastColumn)
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Lấy thông tin vùng merge chứa cell
    /// </summary>
    public MergeRegion? GetMergeRegion(int rowIndex, int columnIndex)
    {
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow <= rowIndex && rowIndex <= region.LastRow &&
                region.FirstColumn <= columnIndex && columnIndex <= region.LastColumn)
            {
                return new MergeRegion(region.FirstRow, region.LastRow, 
                                     region.FirstColumn, region.LastColumn);
            }
        }
        return null;
    }

    /// <summary>
    /// Lấy tất cả các vùng merge trong sheet
    /// </summary>
    public List<MergeRegion> GetAllMergeRegions()
    {
        var regions = new List<MergeRegion>();
        
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            regions.Add(new MergeRegion(region.FirstRow, region.LastRow, 
                                      region.FirstColumn, region.LastColumn));
        }
        
        return regions;
    }

    /// <summary>
    /// Xóa tất cả các vùng merge
    /// </summary>
    public void ClearAllMergeRegions()
    {
        while (_sheet.NumMergedRegions > 0)
        {
            _sheet.RemoveMergedRegion(0);
        }
    }

    /// <summary>
    /// Lấy tất cả merge regions trong một range
    /// </summary>
    public List<CellRangeAddress> GetMergeRegionsInRange(int startRow, int endRow, int startCol, int endCol)
    {
        var regions = new List<CellRangeAddress>();
        
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            
            // Kiểm tra nếu region nằm trong range
            if (region.FirstRow >= startRow && region.LastRow <= endRow &&
                region.FirstColumn >= startCol && region.LastColumn <= endCol)
            {
                regions.Add(region);
            }
        }
        
        return regions;
    }

    /// <summary>
    /// Copy merge regions từ source range đến target range
    /// </summary>
    public void CopyMergeRegions(int sourceStartRow, int sourceEndRow, int sourceStartCol, int sourceEndCol,
                                int targetStartRow, int targetStartCol)
    {
        var sourceMergeRegions = GetMergeRegionsInRange(sourceStartRow, sourceEndRow, sourceStartCol, sourceEndCol);
        
        var rowOffset = targetStartRow - sourceStartRow;
        var colOffset = targetStartCol - sourceStartCol;
        
        foreach (var sourceRegion in sourceMergeRegions)
        {
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
    /// Shift merge regions khi insert/delete rows
    /// </summary>
    public void ShiftMergeRegions(int startRow, int rowShift)
    {
        var regionsToUpdate = new List<(int index, CellRangeAddress region)>();
        
        // Thu thập các regions cần update
        for (int i = 0; i < _sheet.NumMergedRegions; i++)
        {
            var region = _sheet.GetMergedRegion(i);
            if (region.FirstRow >= startRow)
            {
                regionsToUpdate.Add((i, region));
            }
        }

        // Update theo thứ tự ngược để tránh conflict index
        foreach (var (index, region) in regionsToUpdate.OrderByDescending(x => x.index))
        {
            _sheet.RemoveMergedRegion(index);
            
            var newRegion = new CellRangeAddress(
                region.FirstRow + rowShift,
                region.LastRow + rowShift,
                region.FirstColumn,
                region.LastColumn);

            if (newRegion.FirstRow >= 0 && newRegion.LastRow >= 0)
            {
                try
                {
                    _sheet.AddMergedRegion(newRegion);
                }
                catch (ArgumentException)
                {
                    // Invalid region, ignore
                }
            }
        }
    }

    /// <summary>
    /// Khôi phục merge regions với offset cho row và column
    /// </summary>
    public void RestoreMergeRegions(List<MergeRegion> mergeRegions, int rowOffset = 0, int colOffset = 0)
    {
        foreach (var mergeRegion in mergeRegions)
        {
            try
            {
                MergeCells(
                    mergeRegion.FirstRow + rowOffset, 
                    mergeRegion.LastRow + rowOffset,
                    mergeRegion.FirstColumn + colOffset, 
                    mergeRegion.LastColumn + colOffset);
            }
            catch (ArgumentException)
            {
                // Merge region already exists or invalid, ignore
            }
        }
    }
}
