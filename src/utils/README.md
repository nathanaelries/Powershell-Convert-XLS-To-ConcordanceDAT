# Utilities Modules

This directory contains utility modules for Excel handling and file conversion operations.

## Files

### excel-handler.ps1
Handles Excel-specific operations:
- Creating Excel instances
- Proper Excel cleanup
- COM object management

### file-converter.ps1
Manages file conversion operations:
- Converting Excel to temporary format
- Converting temporary files to DAT format
- Line formatting and processing

## Usage

```powershell
using module '.\excel-handler.ps1'
using module '.\file-converter.ps1'

# Excel operations
$excel = New-ExcelInstance
Quit-Excel $excel

# File conversion
Convert-ExcelToTemp -WorksheetPath $path -OutputPath $output
Convert-TempToDAT -InputPath $input -OutputPath $output
```