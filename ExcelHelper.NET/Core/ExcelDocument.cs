namespace ExcelHelper.NET.Core;

/// <summary>
/// Class chính để làm việc với Excel document
/// </summary>
public class ExcelDocument : IDisposable
{
    private XSSFWorkbook? _workbook;
    private readonly string _templatePath;
    private readonly string _outputPath;
    private bool _disposed;

    public ExcelDocument(string templatePath, string outputPath)
    {
        _templatePath = templatePath ?? throw new ArgumentNullException(nameof(templatePath));
        _outputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
        
        if (!File.Exists(_templatePath))
            throw new FileNotFoundException($"Template file không tồn tại: {_templatePath}");
            
        _workbook = LoadWorkbook();
    }

    /// <summary>
    /// Workbook hiện tại
    /// </summary>
    public XSSFWorkbook Workbook => _workbook ?? throw new ObjectDisposedException(nameof(ExcelDocument));

    /// <summary>
    /// Lấy sheet theo index
    /// </summary>
    public ISheet GetSheet(int index)
    {
        if (_workbook == null) throw new ObjectDisposedException(nameof(ExcelDocument));
        
        if (index < 0 || index >= _workbook.NumberOfSheets)
            throw new ArgumentOutOfRangeException(nameof(index), "Sheet index không hợp lệ");
            
        return _workbook.GetSheetAt(index);
    }

    /// <summary>
    /// Lấy sheet theo tên
    /// </summary>
    public ISheet GetSheet(string name)
    {
        if (_workbook == null) throw new ObjectDisposedException(nameof(ExcelDocument));
        
        var sheet = _workbook.GetSheet(name);
        if (sheet == null)
            throw new ArgumentException($"Không tìm thấy sheet: {name}");
            
        return sheet;
    }

    /// <summary>
    /// Tạo SheetManager cho sheet theo index
    /// </summary>
    public SheetManager GetSheetManager(int index)
    {
        var sheet = GetSheet(index);
        return new SheetManager(sheet);
    }

    /// <summary>
    /// Tạo SheetManager cho sheet theo tên
    /// </summary>
    public SheetManager GetSheetManager(string name)
    {
        var sheet = GetSheet(name);
        return new SheetManager(sheet);
    }

    /// <summary>
    /// Lưu file Excel
    /// </summary>
    public string Save(string? password = null)
    {
        if (_workbook == null) throw new ObjectDisposedException(nameof(ExcelDocument));

        try
        {
            // Tạo thư mục output nếu chưa tồn tại
            var outputDir = Path.GetDirectoryName(_outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Xóa file cũ nếu tồn tại
            if (File.Exists(_outputPath))
            {
                File.Delete(_outputPath);
            }

            // Đặt password cho các sheet nếu được yêu cầu
            if (!string.IsNullOrEmpty(password))
            {
                for (int i = 0; i < _workbook.NumberOfSheets; i++)
                {
                    var sheet = _workbook.GetSheetAt(i);
                    if (sheet is XSSFSheet xssfSheet)
                    {
                        xssfSheet.ProtectSheet(password);
                    }
                }
            }

            // Lưu file
            using var fileStream = new FileStream(_outputPath, FileMode.Create, FileAccess.Write);
            _workbook.Write(fileStream);
            
            return _outputPath;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Không thể lưu file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Load workbook từ file template
    /// </summary>
    private XSSFWorkbook LoadWorkbook()
    {
        try
        {
            using var fileStream = new FileStream(_templatePath, FileMode.Open, FileAccess.Read);
            return new XSSFWorkbook(fileStream);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Không thể load file template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Dispose resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed && disposing)
        {
            try
            {
                _workbook?.Close();
            }
            catch
            {
                // Ignore cleanup errors
            }
            finally
            {
                _workbook = null;
                _disposed = true;
            }
        }
    }

    /// <summary>
    /// Finalizer
    /// </summary>
    ~ExcelDocument()
    {
        Dispose(false);
    }
}
