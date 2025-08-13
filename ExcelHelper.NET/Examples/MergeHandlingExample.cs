using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelHelper.NET.Core;
using ExcelHelper.NET.Data;
using System.Data;

namespace ExcelHelper.NET.Examples;

/// <summary>
/// Ví dụ sử dụng các tính năng xử lý merge cells
/// </summary>
public class MergeHandlingExample
{
    public void DemonstrateRangeWithMerge()
    {
        // Tạo workbook mới
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("MergeDemo");
        var sheetManager = new SheetManager(sheet);

        // Tạo template range với merge cells
        CreateTemplateWithMerge(sheet);

        // Test data
        var testData = new List<SampleData>
        {
            new SampleData { Name = "Nguyễn Văn A", Age = 25, Department = "IT", Salary = 15000000 },
            new SampleData { Name = "Trần Thị B", Age = 28, Department = "HR", Salary = 12000000 },
            new SampleData { Name = "Lê Văn C", Age = 30, Department = "Finance", Salary = 18000000 }
        };

        // Chèn dữ liệu vào range với preserve merge
        sheetManager.InsertDataToRange(testData, "A1:D5", null, preserveMerge: true);

        // Lưu file để kiểm tra
        using var fileStream = new FileStream("MergeDemo.xlsx", FileMode.Create);
        workbook.Write(fileStream);
        workbook.Close();

        Console.WriteLine("Demo merge handling completed. Check MergeDemo.xlsx file.");
    }

    public void DemonstrateRowOperationsWithMerge()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("RowMergeDemo");
        var sheetManager = new SheetManager(sheet);

        // Tạo template row với merge
        CreateRowTemplateWithMerge(sheet);

        // Tạo nhiều rows từ template với merge
        sheetManager.CreateRowsWithMerge(templateRowIndex: 0, count: 5);

        // Move row với merge
        sheetManager.MoveRowWithMerge(sourceRowIndex: 2, targetRowIndex: 6);

        // Lưu file
        using var fileStream = new FileStream("RowMergeDemo.xlsx", FileMode.Create);
        workbook.Write(fileStream);
        workbook.Close();

        Console.WriteLine("Demo row merge handling completed. Check RowMergeDemo.xlsx file.");
    }

    public void DemonstrateCopyRangeWithMerge()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("CopyMergeDemo");
        var sheetManager = new SheetManager(sheet);

        // Tạo source range với merge
        CreateComplexRangeWithMerge(sheet);

        // Copy range với merge cells
        sheetManager.CopyRangeWithMerge("A1:E10", targetStartRow: 15);

        // Lưu file
        using var fileStream = new FileStream("CopyMergeDemo.xlsx", FileMode.Create);
        workbook.Write(fileStream);
        workbook.Close();

        Console.WriteLine("Demo copy range with merge completed. Check CopyMergeDemo.xlsx file.");
    }

    public void DemonstrateDataTableWithMerge()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("DataTableMergeDemo");
        var sheetManager = new SheetManager(sheet);

        // Tạo template range với merge cells
        CreateTemplateWithMerge(sheet);

        // Tạo DataTable với sample data
        var dataTable = CreateSampleDataTable();

        // Insert DataTable với merge support - sử dụng generic method
        sheetManager.InsertDataToRange<SampleData>(dataTable, "A1:D5", preserveMerge: true);

        // Lưu file
        using var fileStream = new FileStream("DataTableMergeDemo.xlsx", FileMode.Create);
        workbook.Write(fileStream);
        workbook.Close();

        Console.WriteLine("Demo DataTable merge handling completed. Check DataTableMergeDemo.xlsx file.");
    }

    private void CreateTemplateWithMerge(ISheet sheet)
    {
        // Tạo header với merge cells
        var row1 = sheet.CreateRow(0);
        row1.CreateCell(0).SetCellValue("Employee Information");
        row1.CreateCell(1).SetCellValue("");
        row1.CreateCell(2).SetCellValue("Financial Info");
        row1.CreateCell(3).SetCellValue("");

        // Merge header cells
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1)); // A1:B1
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 3)); // C1:D1

        // Sub headers
        var row2 = sheet.CreateRow(1);
        row2.CreateCell(0).SetCellValue("Name");
        row2.CreateCell(1).SetCellValue("Age");
        row2.CreateCell(2).SetCellValue("Department");
        row2.CreateCell(3).SetCellValue("Salary");

        // Data template rows
        for (int i = 2; i <= 4; i++)
        {
            var row = sheet.CreateRow(i);
            row.CreateCell(0).SetCellValue($"[Name]");
            row.CreateCell(1).SetCellValue($"[Age]");
            row.CreateCell(2).SetCellValue($"[Department]");
            row.CreateCell(3).SetCellValue($"[Salary]");
        }

        // Merge some cells in data area
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 0, 0)); // A3:A4 - Name spanning 2 rows
    }

    private void CreateRowTemplateWithMerge(ISheet sheet)
    {
        var row = sheet.CreateRow(0);
        row.CreateCell(0).SetCellValue("Item");
        row.CreateCell(1).SetCellValue("Description");
        row.CreateCell(2).SetCellValue("Price");
        row.CreateCell(3).SetCellValue("Quantity");
        row.CreateCell(4).SetCellValue("Total");

        // Merge description across 2 columns
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 1, 2)); // B1:C1
    }

    private void CreateComplexRangeWithMerge(ISheet sheet)
    {
        // Tạo một range phức tạp với nhiều merge areas
        for (int row = 0; row < 10; row++)
        {
            var sheetRow = sheet.CreateRow(row);
            for (int col = 0; col < 5; col++)
            {
                sheetRow.CreateCell(col).SetCellValue($"R{row + 1}C{col + 1}");
            }
        }

        // Tạo các merge regions
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 1)); // A1:B2
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 4)); // C1:E1
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 5, 1, 3)); // B4:D6
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(7, 9, 0, 0)); // A8:A10
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(8, 9, 2, 4)); // C9:E10
    }

    /// <summary>
    /// Tạo sample DataTable cho testing
    /// </summary>
    private DataTable CreateSampleDataTable()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Columns.Add("Department", typeof(string));
        dataTable.Columns.Add("Salary", typeof(decimal));

        dataTable.Rows.Add("Nguyễn Văn A", 25, "IT", 15000000);
        dataTable.Rows.Add("Trần Thị B", 28, "HR", 12000000);
        dataTable.Rows.Add("Lê Văn C", 30, "Finance", 18000000);
        dataTable.Rows.Add("Phạm Văn D", 32, "Marketing", 16000000);

        return dataTable;
    }
}

/// <summary>
/// Sample data class for testing
/// </summary>
public class SampleData
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
    public decimal Salary { get; set; }
}
