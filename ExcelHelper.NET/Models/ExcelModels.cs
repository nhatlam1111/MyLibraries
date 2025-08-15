using NPOI.SS.UserModel;
using System.Drawing;

namespace ExcelHelper.NET.Models;

/// <summary>
/// Thông tin cơ bản về cell
/// </summary>
public record CellInfo(int RowIndex, int ColumnIndex, string Address);

/// <summary>
/// Cấu hình định dạng cho cell
/// </summary>
public record CellFormat(
    string? StringFormat = null,
    HorizontalAlignment? HorizontalAlign = null,
    VerticalAlignment? VerticalAlign = null,
    bool HasBorder = false,
    Color? BackgroundColor = null,
    bool IsBold = false,
    bool IsItalic = false,
    bool IsUnderline = false,
    string? FontName = null,
    float? FontSize = null
);

/// <summary>
/// Cấu hình cho việc chèn dữ liệu
/// </summary>
public record InsertOptions(
    bool AutofitHeight = false,
    bool InsertSequence = true,
    int SequenceStartNumber = 1
);

/// <summary>
/// Thông tin về vùng merge
/// </summary>
public record MergeRegion(int FirstRow, int LastRow, int FirstColumn, int LastColumn);

/// <summary>
/// Cấu hình cho việc resize hình ảnh
/// </summary>
public record ImageResizeOptions(
    double? MaxWidth = null,
    double? MaxHeight = null,
    bool MaintainAspectRatio = true
);
