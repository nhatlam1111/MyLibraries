namespace ExcelHelper.NET.Utils;

/// <summary>
/// Tiện ích chuyển đổi giữa đơn vị pixel và đơn vị Excel
/// </summary>
public static class PixelConverter
{
    public const short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    public const short EXCEL_ROW_HEIGHT_FACTOR = 20;
    public const int UNIT_OFFSET_LENGTH = 7;
    public static readonly short[] UNIT_OFFSET_MAP = { 0, 36, 73, 109, 146, 182, 219 };

    /// <summary>
    /// Chuyển pixel thành width units của Excel
    /// </summary>
    public static short PixelToWidthUnits(int pixels)
    {
        short widthUnits = (short)(EXCEL_COLUMN_WIDTH_FACTOR * (pixels / UNIT_OFFSET_LENGTH));
        widthUnits += UNIT_OFFSET_MAP[pixels % UNIT_OFFSET_LENGTH];
        return widthUnits;
    }

    /// <summary>
    /// Chuyển width units của Excel thành pixel
    /// </summary>
    public static int WidthUnitsToPixel(short widthUnits)
    {
        double pixels = (widthUnits / (double)EXCEL_COLUMN_WIDTH_FACTOR) * UNIT_OFFSET_LENGTH;
        int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
        pixels += Math.Floor(offsetWidthUnits / (EXCEL_COLUMN_WIDTH_FACTOR / (double)UNIT_OFFSET_LENGTH));
        return (int)pixels;
    }

    /// <summary>
    /// Chuyển height units của Excel thành pixel
    /// </summary>
    public static int HeightUnitsToPixel(short heightUnits)
    {
        double pixels = heightUnits / (double)EXCEL_ROW_HEIGHT_FACTOR;
        int offsetWidthUnits = heightUnits % EXCEL_ROW_HEIGHT_FACTOR;
        pixels += Math.Floor(offsetWidthUnits / (EXCEL_ROW_HEIGHT_FACTOR / (double)UNIT_OFFSET_LENGTH));
        return (int)pixels;
    }

    /// <summary>
    /// Tính toán scale để resize
    /// </summary>
    public static double CalculateScale(double sourceSize, double targetSize)
    {
        return targetSize / (sourceSize == 0 ? 1 : sourceSize);
    }
}
