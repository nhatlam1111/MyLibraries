using ExcelHelper.NET.Core;
using ExcelHelper.NET.Extensions;
using ExcelHelper.NET.Models;
using NPOI.XSSF.UserModel;

namespace ExcelHelper.NET.Examples;

/// <summary>
/// Quick test để kiểm tra library hoạt động
/// </summary>
public static class QuickTest
{
    public static void RunTest()
    {
        // Test dữ liệu
        var data = new List<Person>
        {
            new("John Doe", 25, "Engineer"),
            new("Jane Smith", 30, "Manager"),
            new("Bob Johnson", 35, "Developer")
        };

        // Tạo empty workbook bằng cách tạo file template trống
        var templatePath = Path.Combine(Path.GetTempPath(), "temp_template.xlsx");
        var outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ExcelHelper_Test.xlsx");
        
        // Tạo template trống
        using (var tempWorkbook = new XSSFWorkbook())
        {
            using var fs = File.Create(templatePath);
            tempWorkbook.Write(fs);
        }

        using var excel = new ExcelDocument(templatePath, outputPath);
        var sheet = excel.Workbook.CreateSheet("Test Sheet");
        
        // Thêm headers
        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("Name");
        headerRow.CreateCell(1).SetCellValue("Age");
        headerRow.CreateCell(2).SetCellValue("Job");
        
        // Style headers
        var headerStyle = excel.Workbook.CreateCellStyle();
        var font = excel.Workbook.CreateFont();
        font.IsBold = true;
        headerStyle.SetFont(font);
        
        for (int i = 0; i < 3; i++)
        {
            headerRow.GetCell(i).CellStyle = headerStyle;
        }

        // Thêm dữ liệu
        for (int i = 0; i < data.Count; i++)
        {
            var row = sheet.CreateRow(i + 1);
            row.CreateCell(0).SetCellValue(data[i].Name);
            row.CreateCell(1).SetCellValue(data[i].Age);
            row.CreateCell(2).SetCellValue(data[i].Job);
        }

        // Auto size columns
        for (int i = 0; i < 3; i++)
        {
            sheet.AutoSizeColumn(i);
        }

        // Save file
        excel.Save();
        
        // Cleanup
        File.Delete(templatePath);
        
        Console.WriteLine($"Test file created: {outputPath}");
        Console.WriteLine($"Sheet has {sheet.LastRowNum + 1} rows");
    }
}

/// <summary>
/// Test model
/// </summary>
public record Person(string Name, int Age, string Job);
