using System.Reflection;

namespace ExcelHelper.NET.Data;

/// <summary>
/// Generic data provider sử dụng reflection để làm việc với bất kỳ kiểu T nào
/// </summary>
/// <typeparam name="T">Kiểu dữ liệu</typeparam>
public class GenericDataProvider<T> : IExcelDataProvider<T>
{
    private readonly Dictionary<string, PropertyInfo> _properties;
    private readonly HashSet<string> _numericFields;
    private readonly HashSet<string> _dateFields;
    private readonly HashSet<string> _imageFields;
    private readonly HashSet<string> _excludedFields;

    public GenericDataProvider(
        HashSet<string>? numericFields = null,
        HashSet<string>? dateFields = null,
        HashSet<string>? imageFields = null,
        HashSet<string>? excludedFields = null)
    {
        _properties = typeof(T).GetProperties()
            .ToDictionary(p => p.Name.ToLower(), p => p);

        _numericFields = numericFields?.Select(f => f.ToLower()).ToHashSet() ?? new HashSet<string>();
        _dateFields = dateFields?.Select(f => f.ToLower()).ToHashSet() ?? new HashSet<string>();
        _imageFields = imageFields?.Select(f => f.ToLower()).ToHashSet() ?? new HashSet<string>();
        _excludedFields = excludedFields?.Select(f => f.ToLower()).ToHashSet() ?? new HashSet<string>();
    }

    public IEnumerable<KeyValuePair<string, object?>> GetFieldValues(T item)
    {
        return _properties
            .Where(p => !_excludedFields.Contains(p.Key))
            .Select(p => new KeyValuePair<string, object?>(p.Key, p.Value.GetValue(item)));
    }

    public bool IsNumericField(string fieldName)
    {
        var lowerName = fieldName.ToLower();
        if (_numericFields.Contains(lowerName))
            return true;

        if (_properties.TryGetValue(lowerName, out var prop))
        {
            var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
            return type == typeof(int) || type == typeof(long) || type == typeof(decimal) 
                || type == typeof(double) || type == typeof(float);
        }

        return false;
    }

    public bool IsDateField(string fieldName)
    {
        var lowerName = fieldName.ToLower();
        if (_dateFields.Contains(lowerName))
            return true;

        if (_properties.TryGetValue(lowerName, out var prop))
        {
            var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
            return type == typeof(DateTime) || type == typeof(DateOnly) || type == typeof(DateTimeOffset);
        }

        return false;
    }

    public bool IsImageField(string fieldName)
    {
        var lowerName = fieldName.ToLower();
        if (_imageFields.Contains(lowerName))
            return true;

        if (_properties.TryGetValue(lowerName, out var prop))
        {
            return prop.PropertyType == typeof(byte[]);
        }

        return false;
    }

    public object? GetFieldValue(T item, string fieldName)
    {
        var lowerName = fieldName.ToLower();
        if (_properties.TryGetValue(lowerName, out var prop))
        {
            return prop.GetValue(item);
        }
        return null;
    }

    public IEnumerable<string> GetFieldNames()
    {
        return _properties.Keys.Where(k => !_excludedFields.Contains(k));
    }
}
