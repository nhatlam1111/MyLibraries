using System.Data;
using System.Reflection;

namespace ExcelHelper.NET.Data;

/// <summary>
/// Adapter để chuyển đổi từ DataTable sang List&lt;T&gt;
/// </summary>
public static class DataTableAdapter
{
    /// <summary>
    /// Chuyển đổi DataTable thành List&lt;T&gt;
    /// </summary>
    /// <typeparam name="T">Kiểu dữ liệu target</typeparam>
    /// <param name="dataTable">DataTable nguồn</param>
    /// <returns>List&lt;T&gt;</returns>
    public static List<T> ToList<T>(DataTable dataTable) where T : new()
    {
        var result = new List<T>();
        var properties = typeof(T).GetProperties()
            .ToDictionary(p => p.Name.ToLower(), p => p);

        foreach (DataRow row in dataTable.Rows)
        {
            var item = new T();
            foreach (DataColumn column in dataTable.Columns)
            {
                if (properties.TryGetValue(column.ColumnName.ToLower(), out var prop))
                {
                    var value = row[column];
                    if (value != DBNull.Value)
                    {
                        try
                        {
                            var convertedValue = ConvertValue(value, prop.PropertyType);
                            prop.SetValue(item, convertedValue);
                        }
                        catch
                        {
                            // Ignore conversion errors
                        }
                    }
                }
            }
            result.Add(item);
        }
        return result;
    }

    /// <summary>
    /// Chuyển đổi giá trị sang kiểu dữ liệu phù hợp
    /// </summary>
    private static object? ConvertValue(object value, Type targetType)
    {
        if (value == null || value == DBNull.Value)
            return null;

        var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

        if (underlyingType == typeof(string))
            return value.ToString();

        if (underlyingType == typeof(DateTime))
        {
            if (DateTime.TryParse(value.ToString(), out var dateTime))
                return dateTime;
        }

        if (underlyingType == typeof(int))
        {
            if (int.TryParse(value.ToString(), out var intValue))
                return intValue;
        }

        if (underlyingType == typeof(decimal))
        {
            if (decimal.TryParse(value.ToString(), out var decimalValue))
                return decimalValue;
        }

        if (underlyingType == typeof(double))
        {
            if (double.TryParse(value.ToString(), out var doubleValue))
                return doubleValue;
        }

        if (underlyingType == typeof(bool))
        {
            if (bool.TryParse(value.ToString(), out var boolValue))
                return boolValue;
        }

        if (underlyingType == typeof(byte[]) && value is byte[] bytes)
            return bytes;

        // Fallback to Convert.ChangeType
        try
        {
            return Convert.ChangeType(value, underlyingType);
        }
        catch
        {
            return null;
        }
    }
}
