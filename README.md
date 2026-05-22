# xls2dat

Cross-platform CLI that converts spreadsheets to [Concordance](https://en.wikipedia.org/wiki/Concordance_(software)) `.dat` files.

Supported inputs:

| Format            | Extensions          | Reader            |
| ----------------- | ------------------- | ----------------- |
| Delimited text    | `.csv`, `.tsv`, `.txt` | native (CsvHelper) |
| Modern Excel      | `.xlsx`, `.xlsm`    | native (ClosedXML) |
| Legacy Excel      | `.xls`              | LibreOffice (auto) |
| LibreOffice Calc  | `.ods`              | LibreOffice (auto) |
| Apple Numbers     | `.numbers`          | LibreOffice (auto) |

Output is Concordance DAT — UTF-16 LE with BOM, field delimiter `0x14`, text qualifier `0xFE` (þ), embedded newlines mapped to `0xAE` (®), CRLF line terminators. All of these are configurable.

## Install

Requires the .NET 8 SDK to build, or download a self-contained binary from the [Releases](../../releases) page (when published).

```sh
git clone https://github.com/nathanaelries/Powershell-Convert-XLS-To-ConcordanceDAT.git
cd Powershell-Convert-XLS-To-ConcordanceDAT
dotnet build -c Release
```

To publish a single-file binary:

```sh
# Windows x64
dotnet publish src/Xls2Dat.Cli -c Release -r win-x64 --self-contained

# Linux x64
dotnet publish src/Xls2Dat.Cli -c Release -r linux-x64 --self-contained

# macOS arm64
dotnet publish src/Xls2Dat.Cli -c Release -r osx-arm64 --self-contained
```

For `.xls`, `.ods`, and `.numbers` inputs, install [LibreOffice](https://www.libreoffice.org/) and make sure `soffice` is on `PATH` (or set `SOFFICE_PATH`, or pass `--soffice-path`).

## Usage

```
xls2dat <input> --output <dir> [options]
```

| Option | Default | Description |
| ------ | ------- | ----------- |
| `--output`, `-o` | _required_ | Directory where `.dat` files will be written (one per worksheet). |
| `--field-delimiter` | `0x14` | Field delimiter byte. Accepts a single char or `0xNN`. |
| `--text-qualifier` | `0xFE` | Text qualifier (wrapping) byte. |
| `--newline-replacement` | `0xAE` | Replaces embedded LFs inside fields. |
| `--encoding` | `utf-16le` | Output encoding: `utf-16le`, `utf-16be`, `utf-8`, `utf-8-bom`. |
| `--skip-header` | _off_ | Drops the first row of each worksheet. |
| `--csv-delimiter` | _inferred_ | Override CSV/TSV delimiter (e.g. `;`). |
| `--soffice-path` | _auto-detect_ | Explicit LibreOffice executable path. |
| `--verbose`, `-v` | _off_ | Verbose log to stderr. |

### Examples

```sh
xls2dat data.xlsx -o ./out
xls2dat report.csv -o ./out --csv-delimiter ';' --skip-header
xls2dat archive.xls -o ./out --soffice-path "/Applications/LibreOffice.app/Contents/MacOS/soffice"
xls2dat tabular.ods -o ./out --encoding utf-8-bom --verbose
```

Each worksheet becomes `<workbook-stem>_<sheet-name>.dat` in the output directory.

## Project layout

```
.
├── src/
│   ├── Xls2Dat.Core/     # library: readers, writer, formatter, detector
│   └── Xls2Dat.Cli/      # net8.0 console (single-file publishable)
├── tests/
│   └── Xls2Dat.Tests/    # xunit + FluentAssertions
├── Xls2Dat.sln
├── TODO.md               # backlog
└── README.md
```

## Tests

```sh
dotnet test
```

## License

MIT. See `LICENSE`.
