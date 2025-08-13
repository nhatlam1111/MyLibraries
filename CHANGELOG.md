# Changelog

All notable changes to MyLibraries will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-08-13

### Added
- **ExcelHelper.NET**: Complete refactor of legacy ExcelHelper.txt
- **SheetManager**: High-level sheet operations with clean API
- **RangeManager**: Advanced range manipulation with merge cell support
- **RowManager**: Row operations (create, clone, move) with merge handling
- **MergeManager**: Comprehensive merge regions management
- **ImageManager**: Image insertion capabilities
- **CellStyler**: Flexible cell formatting and styling system
- **Generic Data Support**: `IExcelDataProvider<T>` for type-safe operations
- **DataTable Integration**: Seamless DataTable to generic type conversion
- **Template Variables**: `$[FieldName]` replacement system
- **Merge Cell Preservation**: Advanced merge cell handling during copy/move operations
- **Examples**: Comprehensive usage examples and demos
- **TestRunner**: Console application for testing

### Architecture
- Modular design with separation of concerns
- Generic type support with `IExcelDataProvider<T>`
- Flexible API with optional parameters
- DRY principle implementation
- SOLID principles adherence

### Features
- ✅ Insert data from `List<T>` with merge preservation
- ✅ Insert data from `DataTable` with automatic type conversion
- ✅ Copy ranges with merge cell handling
- ✅ Create rows from templates with merge support
- ✅ Move rows with merge region updates
- ✅ Template variable replacement system
- ✅ Image insertion with size management
- ✅ Cell formatting and styling
- ✅ Range operations with merge awareness

### Improvements over Legacy Version
- **Architecture**: Monolithic → Modular components
- **Data Types**: DataTable only → Generic + DataTable
- **Merge Handling**: Limited → Full support with flexibility
- **Code Quality**: Duplicate logic → DRY principle
- **API Design**: Complex methods → Clean, intuitive interface
- **Maintainability**: Difficult → Easy with clear separation
- **Testing**: None → Comprehensive examples and tests

### Technical Details
- Target Framework: .NET 8.0 and .NET 10.0
- Dependencies: NPOI (latest)
- Architecture: Clean Architecture principles
- Design Patterns: Factory, Strategy, Adapter

## [Unreleased]
### Planned
- NuGet package creation
- Unit tests with xUnit
- Performance benchmarks
- Additional data providers (JSON, XML)
- Advanced styling templates
- Chart generation support
