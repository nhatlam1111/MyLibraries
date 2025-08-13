using NPOI.SS.UserModel;
using System.Drawing;

namespace ExcelHelper.NET.Extensions;

/// <summary>
/// Extension methods cho ICell
/// </summary>
public static class CellExtensions
{
    /// <summary>
    /// Đặt giá trị cho cell với type checking tự động
    /// </summary>
    public static void SetValue(this ICell cell, object? value, ICellStyle? style = null)
    {
        if (value == null)
        {
            cell.SetBlank();
            if (style != null) cell.CellStyle = style;
            return;
        }

        switch (value)
        {
            case string stringValue:
                cell.SetCellValue(stringValue);
                break;
            case int intValue:
                cell.SetCellValue(intValue);
                break;
            case double doubleValue:
                cell.SetCellValue(doubleValue);
                break;
            case decimal decimalValue:
                cell.SetCellValue((double)decimalValue);
                break;
            case DateTime dateTimeValue:
                cell.SetCellValue(dateTimeValue);
                break;
            case bool boolValue:
                cell.SetCellValue(boolValue);
                break;
            case byte[] _:
                // Xử lý riêng cho hình ảnh - không set value
                cell.SetCellValue("");
                break;
            default:
                // Fallback to string
                cell.SetCellValue(value.ToString() ?? "");
                break;
        }

        // Remove any existing comment
        cell.RemoveCellComment();

        // Apply style if provided
        if (style != null)
            cell.CellStyle = style;
    }

    /// <summary>
    /// Lấy giá trị từ cell dưới dạng string
    /// </summary>
    public static string GetStringValue(this ICell cell)
    {
        try
        {
            return cell.CellType switch
            {
                CellType.Numeric => cell.NumericCellValue.ToString(),
                CellType.String => cell.StringCellValue,
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Formula => cell.StringCellValue,
                _ => ""
            };
        }
        catch
        {
            return "";
        }
    }

    /// <summary>
    /// Kiểm tra cell có rỗng không
    /// </summary>
    public static bool IsEmpty(this ICell? cell)
    {
        if (cell == null) return true;
        
        return cell.CellType switch
        {
            CellType.Blank => true,
            CellType.String => string.IsNullOrWhiteSpace(cell.StringCellValue),
            _ => false
        };
    }
}
