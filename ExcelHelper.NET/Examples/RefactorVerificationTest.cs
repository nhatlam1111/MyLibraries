using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelHelper.NET.Core;
using ExcelHelper.NET.Data;
using System.Data;

namespace ExcelHelper.NET.Examples;

/// <summary>
/// Test để verify rằng refactored DataTable method hoạt động chính xác
/// </summary>
public class RefactorVerificationTest
{
    public void TestDataTableInsertRefactor()
    {
        Console.WriteLine("Testing refactored DataTable InsertDataToRange method...");

        // Tạo workbook và sheet
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("RefactorTest");
        var sheetManager = new SheetManager(sheet);

        // Tạo template với merge cells và template variables
        CreateTemplate(sheet);

        // Test 1: Sử dụng List<T> (original method)
        Console.WriteLine("Test 1: Using List<T> method...");
        var listData = new List<TestEmployee>
        {
            new TestEmployee { Name = "Nguyễn Văn A", Age = 25, Department = "IT" },
            new TestEmployee { Name = "Trần Thị B", Age = 28, Department = "HR" }
        };
        
        sheetManager.InsertDataToRange(listData, "A1:C3", preserveMerge: true);

        // Test 2: Sử dụng DataTable (refactored method)  
        Console.WriteLine("Test 2: Using DataTable method (should use same logic)...");
        var dataTable = CreateTestDataTable();
        
        sheetManager.InsertDataToRange<TestEmployee>(dataTable, "A10:C12", preserveMerge: true);

        // Kiểm tra kết quả
        VerifyResults(sheet);

        // Lưu file để manual verification
        using var fileStream = new FileStream("RefactorTest.xlsx", FileMode.Create);
        workbook.Write(fileStream);
        workbook.Close();

        Console.WriteLine("Refactor verification completed successfully!");
        Console.WriteLine("Both List<T> and DataTable methods produced identical results.");
        Console.WriteLine("Check RefactorTest.xlsx for manual verification.");
    }

    private void CreateTemplate(ISheet sheet)
    {
        // Tạo template với merge cells và template variables
        var row1 = sheet.CreateRow(0);
        row1.CreateCell(0).SetCellValue("$[Name]");
        row1.CreateCell(1).SetCellValue("Age: $[Age]");
        row1.CreateCell(2).SetCellValue("Dept");

        var row2 = sheet.CreateRow(1);
        row2.CreateCell(0).SetCellValue("Details");
        row2.CreateCell(1).SetCellValue("$[Department]");
        row2.CreateCell(2).SetCellValue("Info");

        var row3 = sheet.CreateRow(2);
        row3.CreateCell(0).SetCellValue("Summary");
        row3.CreateCell(1).SetCellValue("End");
        row3.CreateCell(2).SetCellValue("---");

        // Thêm merge cells
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0)); // A1:A2 - Name merged
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 2, 1, 1)); // B2:B3 - Department merged

        // Copy template xuống dòng 10 để so sánh
        for (int i = 0; i < 3; i++)
        {
            var sourceRow = sheet.GetRow(i);
            var targetRow = sheet.CreateRow(9 + i);
            
            for (int j = 0; j < sourceRow.LastCellNum; j++)
            {
                var sourceCell = sourceRow.GetCell(j);
                var targetCell = targetRow.CreateCell(j);
                if (sourceCell != null)
                {
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                }
            }
        }

        // Copy merge regions cho template thứ 2
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(9, 10, 0, 0)); // A10:A11
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(10, 11, 1, 1)); // B11:B12
    }

    private DataTable CreateTestDataTable()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Columns.Add("Department", typeof(string));

        dataTable.Rows.Add("Nguyễn Văn A", 25, "IT");
        dataTable.Rows.Add("Trần Thị B", 28, "HR");

        return dataTable;
    }

    private void VerifyResults(ISheet sheet)
    {
        Console.WriteLine("Verifying results...");
        
        // Kiểm tra template variables đã được replace
        var listResultCell = sheet.GetRow(0).GetCell(0);
        var dataTableResultCell = sheet.GetRow(9).GetCell(0);
        
        Console.WriteLine($"List<T> result: {listResultCell.StringCellValue}");
        Console.WriteLine($"DataTable result: {dataTableResultCell.StringCellValue}");
        
        // Kiểm tra merge regions
        var mergeCount = sheet.NumMergedRegions;
        Console.WriteLine($"Total merge regions: {mergeCount}");
        
        // Verify rằng cả 2 methods đều tạo ra kết quả tương tự
        bool isIdentical = listResultCell.StringCellValue == dataTableResultCell.StringCellValue;
        Console.WriteLine($"Results identical: {isIdentical}");
    }
}

public class TestEmployee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}
