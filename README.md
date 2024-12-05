# Excel to DAT Converter

This PowerShell script converts Excel files to DAT format with specific formatting requirements, optimized for large files.

## Directory Structure

```
.
├── src/
│   ├── config/
│   │   └── settings.ps1      # Configuration settings
│   ├── utils/
│   │   ├── excel-handler.ps1 # Excel operations
│   │   └── file-converter.ps1# File conversion utilities
│   └── main.ps1             # Main script
└── README.md
```

## Usage

```powershell
.\src\main.ps1 -ExcelFile "path\to\excel.xls" -OutputFolder "path\to\output\folder" -BatchSize 10000 -MaxThreads 4
```

Note: Do not include square brackets [] in the command. They indicate optional parameters in documentation only.

### Parameters

- `ExcelFile`: (Required) Path to the Excel file to convert
- `OutputFolder`: (Required) Path where DAT files will be saved
- `BatchSize`: (Optional) Number of rows to process at once (default: 10000)
- `MaxThreads`: (Optional) Number of parallel processing threads (default: 4)

### Example

```powershell
.\src\main.ps1 -ExcelFile "C:\Data\MyFile.xlsx" -OutputFolder "C:\Output" -BatchSize 20000 -MaxThreads 8
```

## Performance Features

- Parallel processing with configurable thread count
- Batch processing for memory efficiency
- Optimized file I/O with buffering
- Pre-compiled regex patterns
- Memory-efficient streaming
- Configurable batch sizes
- Excel optimization flags
- Proper memory management and cleanup

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Excel installed
- Appropriate permissions for file operations

## Performance Tips

1. Adjust BatchSize based on available RAM
2. Set MaxThreads to match CPU cores (but not exceed them)
3. For very large files, increase buffer sizes in settings.ps1
4. Ensure adequate disk space for temporary files

## Error Handling

The script includes comprehensive error handling:
- Validates input parameters
- Checks file existence
- Ensures proper Excel cleanup
- Manages temporary files
- Reports detailed progress and errors
