using ExcelHelper.NET.Core;
using ExcelHelper.NET.Data;
using ExcelHelper.NET.Models;
using ExcelHelper.NET.Extensions;
using NPOI.SS.UserModel;
using System.Drawing;

namespace ExcelHelper.NET.Examples;

/// <summary>
/// Ví dụ sử dụng ExcelHelper.NET
/// </summary>
public class BasicUsageExample
{
    // Sample data model
    public record Customer(
        int Id,
        string Name,
        string Contact,
        decimal Revenue,
        DateTime CreatedDate,
        string Department,
        byte[]? Avatar = null
    );

    public static void RunExample()
    {
        // Tạo dữ liệu mẫu
        var customers = GenerateSampleData();

        // Ví dụ 1: Sử dụng cơ bản
        BasicUsage(customers);

        // Ví dụ 2: Với styling và formatting
        AdvancedUsage(customers);

        // Ví dụ 3: Làm việc với images
        ImageExample(customers);

        // Ví dụ 4: Range operations
        RangeOperationsExample(customers);
    }

    private static void BasicUsage(List<Customer> customers)
    {
        using var excel = new ExcelDocument("template.xlsx", "basic_output.xlsx");
        var sheetManager = excel.GetSheetManager("Sheet1");

        // Cấu hình data provider
        var dataProvider = new GenericDataProvider<Customer>(
            dateFields: new HashSet<string> { "createddate" },
            numericFields: new HashSet<string> { "id", "revenue" },
            excludedFields: new HashSet<string> { "avatar" } // Loại trừ binary data
        );

        // Chèn dữ liệu với tùy chọn
        var options = new InsertOptions(
            AutofitHeight: true,
            InsertSequence: true,
            SequenceStartNumber: 1
        );

        sheetManager.InsertData(customers, startRow: 2, dataProvider, options);
        excel.Save();
        
        Console.WriteLine("✓ Basic usage example completed!");
    }

    private static void AdvancedUsage(List<Customer> customers)
    {
        using var excel = new ExcelDocument("template.xlsx", "advanced_output.xlsx");
        var sheetManager = excel.GetSheetManager("Sheet1");

        // Data provider
        var dataProvider = new GenericDataProvider<Customer>(
            dateFields: new HashSet<string> { "createddate" },
            numericFields: new HashSet<string> { "id", "revenue" }
        );

        // Insert data
        sheetManager.InsertData(customers, startRow: 3, dataProvider);

        // Styling header
        var headerFormat = new CellFormat(
            HasBorder: true,
            IsBold: true,
            HorizontalAlign: HorizontalAlignment.Center,
            BackgroundColor: Color.LightBlue,
            FontName: "Arial",
            FontSize: 12
        );
        sheetManager.ApplyCellFormat("A2:G2", headerFormat);

        // Styling data rows
        var dataFormat = new CellFormat(
            HasBorder: true,
            HorizontalAlign: HorizontalAlignment.Left,
            VerticalAlign: VerticalAlignment.Center
        );
        sheetManager.ApplyCellFormat($"A3:G{2 + customers.Count}", dataFormat);

        // Format currency column
        var currencyFormat = new CellFormat(
            StringFormat: "#,##0.00",
            HorizontalAlign: HorizontalAlignment.Right,
            HasBorder: true
        );
        sheetManager.ApplyCellFormat($"D3:D{2 + customers.Count}", currencyFormat);

        // Set background color for high revenue customers
        for (int i = 0; i < customers.Count; i++)
        {
            if (customers[i].Revenue > 6000000)
            {
                var rowIndex = 3 + i;
                sheetManager.Styles.SetBackgroundColor(sheetManager.Sheet, $"A{rowIndex}:G{rowIndex}", Color.LightGreen);
            }
        }

        excel.Save("advanced_password");
        Console.WriteLine("✓ Advanced usage example completed!");
    }

    private static void ImageExample(List<Customer> customers)
    {
        using var excel = new ExcelDocument("template.xlsx", "image_output.xlsx");
        var sheetManager = excel.GetSheetManager("Sheet1");

        // Tạo sample image data (placeholder)
        var sampleImage = CreateSampleImageData();

        // Chèn image vào cell
        var imageCell = sheetManager.Sheet.GetCellByIndex(1, 0);
        var resizeOptions = new ImageResizeOptions(
            MaxWidth: 100,
            MaxHeight: 80,
            MaintainAspectRatio: true
        );

        sheetManager.Images.InsertImage(imageCell, sampleImage, resizeOptions);

        // Insert data với images
        var customersWithImages = customers.Take(3).Select(c => c with { Avatar = sampleImage }).ToList();
        var dataProvider = new GenericDataProvider<Customer>(
            imageFields: new HashSet<string> { "avatar" }
        );

        sheetManager.InsertData(customersWithImages, startRow: 5, dataProvider);

        excel.Save();
        Console.WriteLine("✓ Image example completed!");
    }

    private static void RangeOperationsExample(List<Customer> customers)
    {
        using var excel = new ExcelDocument("template.xlsx", "range_output.xlsx");
        var sheetManager = excel.GetSheetManager("Sheet1");

        // Copy header range multiple times
        sheetManager.Ranges.CopyRange("A1:G1", targetStartRow: 5);
        sheetManager.Ranges.CopyRange("A1:G1", targetStartRow: 10);

        // Merge title cells
        sheetManager.Merges.MergeRange("A1:G1");

        // Set range values
        sheetManager.Ranges.SetRangeValue("A15:C15", "Summary Data");

        // Create template ranges for repeating data
        var templateCustomers = customers.Take(2).ToList();
        sheetManager.InsertDataToRange(templateCustomers, "A20:G21");

        excel.Save();
        Console.WriteLine("✓ Range operations example completed!");
    }

    private static List<Customer> GenerateSampleData()
    {
        return new List<Customer>
        {
            new Customer(1, "ABC Corporation", "John Doe", 5000000m, DateTime.Now.AddMonths(-3), "Sales"),
            new Customer(2, "XYZ Limited", "Jane Smith", 7500000m, DateTime.Now.AddMonths(-2), "Marketing"),
            new Customer(3, "Tech Solutions", "Bob Johnson", 3200000m, DateTime.Now.AddMonths(-1), "IT"),
            new Customer(4, "Global Industries", "Alice Brown", 8900000m, DateTime.Now.AddDays(-15), "Operations"),
            new Customer(5, "Innovation Labs", "Charlie Wilson", 4100000m, DateTime.Now.AddDays(-5), "R&D")
        };
    }

    private static byte[] CreateSampleImageData()
    {
        // Tạo một hình ảnh đơn giản (placeholder)
        using var bitmap = new Bitmap(50, 50);
        using var graphics = Graphics.FromImage(bitmap);
        
        graphics.FillRectangle(Brushes.Blue, 0, 0, 50, 50);
        graphics.DrawString("IMG", new Font("Arial", 8), Brushes.White, 5, 15);
        
        using var stream = new MemoryStream();
        bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
        return stream.ToArray();
    }
}

/// <summary>
/// Ví dụ migration từ code cũ
/// </summary>
public class MigrationExample
{
    public static void OldWayExample()
    {
        // Code cũ (commented out vì không có class cũ)
        /*
        var helper = new ExcelHelper(templatePath, outputPath, userId);
        helper.excludedColumn.Add("avatar");
        helper.dateColumns.Add("createddate");
        helper.InsertData(sheet, dataTable, startRow, true);
        helper.SetBackgroundColor(sheet, "A1:G1", Color.Blue);
        helper.WriteDoc("password");
        */
    }

    public static void NewWayExample()
    {
        // Code mới
        using var excel = new ExcelDocument("template.xlsx", "output.xlsx");
        var sheetManager = excel.GetSheetManager("Sheet1");
        
        var dataProvider = new GenericDataProvider<Customer>(
            excludedFields: new HashSet<string> { "avatar" },
            dateFields: new HashSet<string> { "createddate" }
        );
        
        var customers = new List<Customer>(); // Your data
        var options = new InsertOptions(AutofitHeight: true);
        
        sheetManager.InsertData(customers, startRow: 2, dataProvider, options);
        sheetManager.Styles.SetBackgroundColor(sheetManager.Sheet, "A1:G1", Color.Blue);
        excel.Save("password");
        
        Console.WriteLine("✓ Migration example completed!");
    }
}

public record Customer(int Id, string Name, string Contact, decimal Revenue, DateTime CreatedDate, string Department, byte[]? Avatar = null);
