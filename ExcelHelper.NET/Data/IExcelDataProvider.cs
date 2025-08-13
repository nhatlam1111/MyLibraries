namespace ExcelHelper.NET.Data;

/// <summary>
/// Interface chung cho nguồn dữ liệu Excel
/// </summary>
/// <typeparam name="T">Kiểu dữ liệu</typeparam>
public interface IExcelDataProvider<T>
{
    /// <summary>
    /// Lấy tất cả các field value của một item
    /// </summary>
    IEnumerable<KeyValuePair<string, object?>> GetFieldValues(T item);
    
    /// <summary>
    /// Kiểm tra field có phải kiểu số không
    /// </summary>
    bool IsNumericField(string fieldName);
    
    /// <summary>
    /// Kiểm tra field có phải kiểu ngày không
    /// </summary>
    bool IsDateField(string fieldName);
    
    /// <summary>
    /// Kiểm tra field có phải hình ảnh không
    /// </summary>
    bool IsImageField(string fieldName);
    
    /// <summary>
    /// Lấy giá trị của một field cụ thể
    /// </summary>
    object? GetFieldValue(T item, string fieldName);
    
    /// <summary>
    /// Lấy danh sách tên các fields
    /// </summary>
    IEnumerable<string> GetFieldNames();
}
