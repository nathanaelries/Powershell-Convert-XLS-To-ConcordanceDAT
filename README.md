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
.\src\main.ps1 -ExcelFile "path\to\excel.xls" -OutputFolder "path\to\output\folder" [-BatchSize 10000] [-MaxThreads 4]
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
2. Set MaxThreads to match CPU cores
3. For very large files, increase buffer sizes in settings.ps1
4. Ensure adequate disk space for temporary files