# MyLibraries

A collection of .NET libraries for various tasks, currently featuring ExcelHelper.NET - a powerful Excel manipulation library.

## Projects

### 📊 ExcelHelper.NET
A modern, refactored Excel helper library built on top of NPOI with improved architecture and merge cell handling capabilities.

**Features:**
- ✅ Generic data support with `IExcelDataProvider<T>`
- ✅ DataTable integration
- ✅ Advanced merge cell handling
- ✅ Template variable system (`$[FieldName]`)
- ✅ Range operations with merge preservation
- ✅ Row operations with merge support
- ✅ Flexible styling system
- ✅ Image insertion capabilities

**Architecture:**
- `SheetManager` - High-level sheet operations
- `RangeManager` - Range manipulation and copying
- `RowManager` - Row creation, cloning, and moving
- `MergeManager` - Merge regions management
- `ImageManager` - Image handling
- `CellStyler` - Cell formatting and styling

### 🧪 TestRunner
A console application for testing the libraries.

## Quick Start

```csharp
// Create SheetManager
var sheetManager = new SheetManager(sheet);

// Insert data with merge preservation
var employees = new List<Employee> { /* data */ };
sheetManager.InsertDataToRange(employees, "A1:D5", preserveMerge: true);

// Copy range with merge cells
sheetManager.CopyRangeWithMerge("A1:E10", targetStartRow: 15);

// Create rows from template with merge
sheetManager.CreateRowsWithMerge(templateRowIndex: 0, count: 5);

// DataTable support
DataTable dataTable = GetDataTable();
sheetManager.InsertDataToRange<Employee>(dataTable, "A1:D8", preserveMerge: true);
```

## Getting Started

### Prerequisites
- .NET 8.0 or later
- Visual Studio 2022 or VS Code

### Building
```bash
dotnet build MyLibraries.sln
```

### Running Tests
```bash
dotnet run --project TestRunner
```

## Documentation

- [Merge Cell Features](ExcelHelper.NET/MERGE_FEATURES.md) - Detailed documentation of merge cell handling
- [Examples](ExcelHelper.NET/Examples/) - Usage examples and demos

## Comparison with Legacy Version

This library is a complete refactor of the original ExcelHelper.txt with the following improvements:

| Feature | Legacy | ExcelHelper.NET |
|---------|--------|-----------------|
| Architecture | Monolithic class | Modular components |
| Merge Handling | Limited, hardcoded | Full support, flexible |
| Data Types | DataTable only | Generic + DataTable |
| API Design | Complex methods | Clean, intuitive |
| Code Reuse | Duplicate logic | DRY principle |
| Maintainability | Difficult | Easy |

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built on top of [NPOI](https://github.com/tonyqus/npoi)
- Inspired by the need for better Excel manipulation in .NET projects

---

**Author:** Lam Nguyen  
**Date:** August 2025
