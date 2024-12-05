# Configuration Module

This directory contains configuration settings for the Excel to DAT converter.

## Files

- `settings.ps1`: Contains global configuration variables like special characters used in the conversion process.

## Usage

```powershell
using module '.\settings.ps1'
$CONFIG.PIPE    # Access pipe character
$CONFIG.THORNE  # Access thorne character
$CONFIG.TAB     # Access tab character
```