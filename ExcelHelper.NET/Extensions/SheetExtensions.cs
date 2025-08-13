using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula;
using ExcelHelper.NET.Models;

namespace ExcelHelper.NET.Extensions;

/// <summary>
/// Extension methods cho ISheet
/// </summary>
public static class SheetExtensions
{
    /// <summary>
    /// Lấy cell theo địa chỉ Excel (A1, B5, ...)
    /// </summary>
    public static ICell? GetCellByAddress(this ISheet sheet, string address)
    {
        try
        {
            var cellReference = new CellReference(address);
            var row = sheet.GetRow(cellReference.Row) ?? sheet.CreateRow(cellReference.Row);
            return row.GetCell(cellReference.Col) ?? row.CreateCell(cellReference.Col);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Lấy hoặc tạo cell theo index
    /// </summary>
    public static ICell GetCellByIndex(this ISheet sheet, int rowIndex, int columnIndex)
    {
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        return row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
    }

    /// <summary>
    /// Tạo nhiều dòng từ một template row
    /// </summary>
    public static void CreateRowsFromTemplate(this ISheet sheet, int templateRowIndex, int count, bool moveExistingRows = true)
    {
        if (count <= 1) return;

        if (moveExistingRows)
        {
            // Di chuyển các dòng hiện có xuống dưới
            var rowsToMove = sheet.LastRowNum - templateRowIndex;
            for (int i = rowsToMove; i >= 1; i--)
            {
                sheet.MoveRow(templateRowIndex + i, templateRowIndex + i + count - 1);
            }
        }

        // Clone template row
        for (int i = 1; i < count; i++)
        {
            sheet.CloneRow(templateRowIndex, templateRowIndex + i);
        }
    }

    /// <summary>
    /// Clone một dòng
    /// </summary>
    public static void CloneRow(this ISheet sheet, int sourceRowIndex, int targetRowIndex)
    {
        var sourceRow = sheet.GetRow(sourceRowIndex);
        if (sourceRow == null) return;

        var targetRow = sheet.GetRow(targetRowIndex) ?? sheet.CreateRow(targetRowIndex);

        // Copy row properties
        targetRow.Height = sourceRow.Height;
        targetRow.Hidden = sourceRow.Hidden;

        // Copy cells
        for (int i = 0; i < sourceRow.LastCellNum; i++)
        {
            var sourceCell = sourceRow.GetCell(i);
            if (sourceCell == null) continue;

            var targetCell = targetRow.GetCell(i) ?? targetRow.CreateCell(i);
            
            // Copy cell value and type
            switch (sourceCell.CellType)
            {
                case CellType.String:
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                case CellType.Numeric:
                    targetCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.Boolean:
                    targetCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Formula:
                    targetCell.SetCellFormula(sourceCell.CellFormula);
                    break;
                case CellType.Error:
                    targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                    break;
                default:
                    targetCell.SetBlank();
                    break;
            }

            // Copy cell style
            targetCell.CellStyle = sourceCell.CellStyle;
        }

        // Copy merged regions
        CopyMergedRegionsForRow(sheet, sourceRowIndex, targetRowIndex);
    }

    /// <summary>
    /// Di chuyển một dòng
    /// </summary>
    public static void MoveRow(this ISheet sheet, int sourceRowIndex, int targetRowIndex)
    {
        sheet.CloneRow(sourceRowIndex, targetRowIndex);
        
        // Clear source row
        var sourceRow = sheet.GetRow(sourceRowIndex);
        if (sourceRow != null)
        {
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                var cell = sourceRow.GetCell(i);
                cell?.SetBlank();
            }
        }
    }

    /// <summary>
    /// Copy merged regions cho một dòng cụ thể
    /// </summary>
    private static void CopyMergedRegionsForRow(ISheet sheet, int sourceRowIndex, int targetRowIndex)
    {
        for (int i = 0; i < sheet.NumMergedRegions; i++)
        {
            var region = sheet.GetMergedRegion(i);
            if (region.FirstRow == sourceRowIndex && region.LastRow == sourceRowIndex)
            {
                var newRegion = new CellRangeAddress(
                    targetRowIndex,
                    targetRowIndex, 
                    region.FirstColumn,
                    region.LastColumn);
                
                try
                {
                    sheet.AddMergedRegion(newRegion);
                }
                catch
                {
                    // Ignore merge conflicts
                }
            }
        }
    }

    /// <summary>
    /// Xóa một cột và cập nhật merged regions
    /// </summary>
    public static void DeleteColumn(this ISheet sheet, int columnIndex)
    {
        // Lưu danh sách các vùng merge
        var mergedRegions = new List<CellRangeAddress>();
        for (int i = 0; i < sheet.NumMergedRegions; i++)
        {
            mergedRegions.Add(sheet.GetMergedRegion(i));
        }

        // Xóa tất cả merged regions
        while (sheet.NumMergedRegions > 0)
        {
            sheet.RemoveMergedRegion(0);
        }

        // Xóa cells trong cột
        for (int i = 0; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            if (row != null)
            {
                var cell = row.GetCell(columnIndex);
                if (cell != null)
                {
                    row.RemoveCell(cell);
                }

                // Shift cells to the left
                for (int j = columnIndex + 1; j < row.LastCellNum; j++)
                {
                    var cellToMove = row.GetCell(j);
                    if (cellToMove != null)
                    {
                        var newCell = row.GetCell(j - 1) ?? row.CreateCell(j - 1);
                        // Copy cell value and style
                        switch (cellToMove.CellType)
                        {
                            case CellType.String:
                                newCell.SetCellValue(cellToMove.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(cellToMove.NumericCellValue);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(cellToMove.BooleanCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(cellToMove.CellFormula);
                                break;
                            default:
                                newCell.SetBlank();
                                break;
                        }
                        newCell.CellStyle = cellToMove.CellStyle;
                        row.RemoveCell(cellToMove);
                    }
                }
            }
        }

        // Tái tạo merged regions với điều chỉnh
        foreach (var region in mergedRegions)
        {
            var firstCol = region.FirstColumn;
            var lastCol = region.LastColumn;
            var firstRow = region.FirstRow;
            var lastRow = region.LastRow;

            if (firstCol > columnIndex)
            {
                // Shift left
                firstCol--;
                lastCol--;
            }
            else if (firstCol <= columnIndex && lastCol >= columnIndex)
            {
                // Region intersects with deleted column
                if (firstCol == lastCol)
                {
                    // Skip this region as it only spans the deleted column
                    continue;
                }
                lastCol--;
            }

            if (firstCol <= lastCol && firstRow <= lastRow)
            {
                try
                {
                    sheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
                }
                catch
                {
                    // Ignore merge conflicts
                }
            }
        }
    }
}
