using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using ExcelHelper.NET.Models;
using ExcelHelper.NET.Extensions;

namespace ExcelHelper.NET.Styling;

/// <summary>
/// Class quản lý style cho cell
/// </summary>
public class CellStyler
{
    private readonly XSSFWorkbook _workbook;

    public CellStyler(XSSFWorkbook workbook)
    {
        _workbook = workbook;
    }

    /// <summary>
    /// Tạo font mới
    /// </summary>
    public IFont CreateFont(string? fontName = null, float? fontSize = null, bool isBold = false, 
        bool isItalic = false, bool isUnderline = false, Color? color = null)
    {
        var font = _workbook.CreateFont();
        
        if (!string.IsNullOrEmpty(fontName))
            font.FontName = fontName;
            
        if (fontSize.HasValue)
            font.FontHeightInPoints = fontSize.Value;
            
        font.IsBold = isBold;
        font.IsItalic = isItalic;
        font.Underline = isUnderline ? FontUnderlineType.Single : FontUnderlineType.None;

        if (color.HasValue)
        {
            if (font is XSSFFont xssfFont)
            {
                xssfFont.SetColor(new XSSFColor(SixLabors.ImageSharp.Color.FromRgb(color.Value.R, color.Value.G, color.Value.B)));
            }
        }

        return font;
    }

    /// <summary>
    /// Tạo cell style từ CellFormat
    /// </summary>
    public ICellStyle CreateStyle(CellFormat format)
    {
        var style = _workbook.CreateCellStyle();

        // Set border
        if (format.HasBorder)
        {
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.RightBorderColor = HSSFColor.Black.Index;
            style.TopBorderColor = HSSFColor.Black.Index;
        }

        // Set alignment
        if (format.HorizontalAlign.HasValue)
            style.Alignment = format.HorizontalAlign.Value;
            
        if (format.VerticalAlign.HasValue)
            style.VerticalAlignment = format.VerticalAlign.Value;

        // Set background color
        if (format.BackgroundColor.HasValue && style is XSSFCellStyle xssfStyle)
        {
            var color = new XSSFColor(SixLabors.ImageSharp.Color.FromRgb(
                format.BackgroundColor.Value.R, 
                format.BackgroundColor.Value.G, 
                format.BackgroundColor.Value.B 
            ));
            xssfStyle.SetFillForegroundColor(color);
            xssfStyle.FillPattern = FillPattern.SolidForeground;
        }

        // Set font
        if (format.FontName != null || format.FontSize.HasValue || format.IsBold || format.IsItalic || format.IsUnderline)
        {
            var font = CreateFont(format.FontName, format.FontSize, format.IsBold, format.IsItalic, format.IsUnderline);
            style.SetFont(font);
        }

        // Set data format
        if (!string.IsNullOrEmpty(format.StringFormat))
        {
            try
            {
                style.DataFormat = _workbook.GetCreationHelper().CreateDataFormat().GetFormat(format.StringFormat);
            }
            catch
            {
                // Fallback to built-in format
                var dataFormat = _workbook.CreateDataFormat();
                style.DataFormat = dataFormat.GetFormat(format.StringFormat);
            }
        }

        return style;
    }

    /// <summary>
    /// Áp dụng màu nền cho cell
    /// </summary>
    public void SetBackgroundColor(ICell cell, Color color)
    {
        if (cell.CellStyle is XSSFCellStyle xssfStyle)
        {
            var newStyle = (XSSFCellStyle)_workbook.CreateCellStyle();
            newStyle.CloneStyleFrom(xssfStyle);
            
            var xssfColor = new XSSFColor(SixLabors.ImageSharp.Color.FromRgb(color.R, color.G, color.B ));
            newStyle.SetFillForegroundColor(xssfColor);
            newStyle.FillPattern = FillPattern.SolidForeground;
            
            cell.CellStyle = newStyle;
        }
    }

    /// <summary>
    /// Áp dụng màu nền cho một vùng cells
    /// </summary>
    public void SetBackgroundColor(ISheet sheet, string range, Color color)
    {
        var cells = range.Split(':');
        if (cells.Length != 2) return;

        var startCell = sheet.GetCellByAddress(cells[0]);
        var endCell = sheet.GetCellByAddress(cells[1]);
        
        if (startCell == null || endCell == null) return;

        for (int row = startCell.RowIndex; row <= endCell.RowIndex; row++)
        {
            for (int col = startCell.ColumnIndex; col <= endCell.ColumnIndex; col++)
            {
                var cell = sheet.GetCellByIndex(row, col);
                SetBackgroundColor(cell, color);
            }
        }
    }

    /// <summary>
    /// Áp dụng font cho cell
    /// </summary>
    public void SetFont(ICell cell, IFont font)
    {
        var newStyle = _workbook.CreateCellStyle();
        newStyle.CloneStyleFrom(cell.CellStyle);
        newStyle.SetFont(font);
        cell.CellStyle = newStyle;
    }
}
