using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using ExcelHelper.NET.Models;
using ExcelHelper.NET.Utils;

namespace ExcelHelper.NET.Media;

/// <summary>
/// Class quản lý hình ảnh trong Excel
/// </summary>
public class ImageManager
{
    private readonly ISheet _sheet;
    private readonly XSSFWorkbook _workbook;

    public ImageManager(ISheet sheet)
    {
        _sheet = sheet;
        _workbook = (XSSFWorkbook)sheet.Workbook;
    }

    /// <summary>
    /// Chèn hình ảnh vào một cell
    /// </summary>
    public void InsertImage(ICell cell, byte[] imageData, ImageResizeOptions? options = null)
    {
        try
        {
            var resizedImage = options != null ? ResizeImage(imageData, options) : imageData;
            
            int pictureIndex = _workbook.AddPicture(resizedImage, XSSFWorkbook.PICTURE_TYPE_JPG);
            var helper = _workbook.GetCreationHelper() as XSSFCreationHelper;
            var drawing = _sheet.CreateDrawingPatriarch() as XSSFDrawing;
            var anchor = helper!.CreateClientAnchor() as XSSFClientAnchor;
            
            anchor!.AnchorType = AnchorType.MoveAndResize;
            anchor.Col1 = cell.ColumnIndex;
            anchor.Row1 = cell.RowIndex;
            
            var picture = drawing!.CreatePicture(anchor, pictureIndex) as XSSFPicture;
            
            // Calculate scale if options provided
            if (options?.MaxWidth.HasValue == true || options?.MaxHeight.HasValue == true)
            {
                var cellWidth = _sheet.GetColumnWidthInPixels(cell.ColumnIndex);
                var cellHeight = _sheet.GetRow(cell.RowIndex)?.HeightInPoints ?? 15;
                
                var targetWidth = options.MaxWidth ?? cellWidth;
                var targetHeight = options.MaxHeight ?? cellHeight;
                
                var scaleX = PixelConverter.CalculateScale(cellWidth, targetWidth);
                var scaleY = PixelConverter.CalculateScale(cellHeight, targetHeight);
                
                picture!.Resize(scaleX, scaleY);
            }
            
            // Clear cell content
            cell.SetCellValue("");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Không thể chèn hình ảnh: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Chèn hình ảnh vào vùng merged cells
    /// </summary>
    public void InsertImageInMergedRegion(ICell startCell, ICell endCell, byte[] imageData)
    {
        try
        {
            int pictureIndex = _workbook.AddPicture(imageData, XSSFWorkbook.PICTURE_TYPE_JPG);
            var helper = _workbook.GetCreationHelper() as XSSFCreationHelper;
            var drawing = _sheet.CreateDrawingPatriarch() as XSSFDrawing;
            var anchor = helper!.CreateClientAnchor() as XSSFClientAnchor;
            
            anchor!.AnchorType = AnchorType.MoveAndResize;
            anchor.Col1 = startCell.ColumnIndex;
            anchor.Row1 = startCell.RowIndex;
            anchor.Col2 = endCell.ColumnIndex + 1;
            anchor.Row2 = endCell.RowIndex + 1;
            
            // Add some padding
            anchor.Dx1 = 10000;
            anchor.Dy1 = 10000;
            
            drawing!.CreatePicture(anchor, pictureIndex);
            
            // Clear start cell content
            startCell.SetCellValue("");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Không thể chèn hình ảnh vào vùng merge: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Chèn hình ảnh từ System.Drawing.Image
    /// </summary>
    public void InsertImage(ICell cell, Image image, ImageResizeOptions? options = null)
    {
        byte[] imageBytes;
        using (var memoryStream = new MemoryStream())
        {
            image.Save(memoryStream, ImageFormat.Png);
            imageBytes = memoryStream.ToArray();
        }
        
        InsertImage(cell, imageBytes, options);
    }

    /// <summary>
    /// Resize hình ảnh theo tùy chọn
    /// </summary>
    private byte[] ResizeImage(byte[] imageData, ImageResizeOptions options)
    {
        try
        {
            using var originalStream = new MemoryStream(imageData);
            using var originalBitmap = new Bitmap(originalStream);
            
            int newWidth = originalBitmap.Width;
            int newHeight = originalBitmap.Height;
            
            if (options.MaxWidth.HasValue || options.MaxHeight.HasValue)
            {
                var maxWidth = options.MaxWidth ?? int.MaxValue;
                var maxHeight = options.MaxHeight ?? int.MaxValue;
                
                if (options.MaintainAspectRatio)
                {
                    var ratioX = maxWidth / originalBitmap.Width;
                    var ratioY = maxHeight / originalBitmap.Height;
                    var ratio = Math.Min(ratioX, ratioY);
                    
                    newWidth = (int)(originalBitmap.Width * ratio);
                    newHeight = (int)(originalBitmap.Height * ratio);
                }
                else
                {
                    newWidth = (int)Math.Min(maxWidth, originalBitmap.Width);
                    newHeight = (int)Math.Min(maxHeight, originalBitmap.Height);
                }
            }
            
            if (newWidth == originalBitmap.Width && newHeight == originalBitmap.Height)
            {
                return imageData; // No resize needed
            }
            
            using var resizedBitmap = new Bitmap(newWidth, newHeight);
            using var graphics = Graphics.FromImage(resizedBitmap);
            
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
            
            using var outputStream = new MemoryStream();
            resizedBitmap.Save(outputStream, ImageFormat.Png);
            return outputStream.ToArray();
        }
        catch
        {
            // Return original if resize fails
            return imageData;
        }
    }
}
