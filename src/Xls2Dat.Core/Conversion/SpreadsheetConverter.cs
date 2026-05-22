using System;
using System.IO;
using System.Linq;
using Xls2Dat.Core.Detection;
using Xls2Dat.Core.Formatting;
using Xls2Dat.Core.Readers;
using Xls2Dat.Core.Writers;

namespace Xls2Dat.Core.Conversion
{
    public sealed class SpreadsheetConverter
    {
        private readonly ConversionOptions _options;
        private readonly Action<string>? _log;

        public SpreadsheetConverter(ConversionOptions options, Action<string>? log = null)
        {
            _options = options;
            _log = log;
        }

        public ConversionResult Convert(string inputPath, string outputDirectory)
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input file not found.", inputPath);

            Directory.CreateDirectory(outputDirectory);

            var format = FormatDetector.Detect(inputPath);
            _log?.Invoke($"Detected format: {format} ({Path.GetFileName(inputPath)})");

            using var reader = OpenReader(inputPath, format);
            var formatter = new DatLineFormatter(
                _options.FieldDelimiter,
                _options.TextQualifier,
                _options.NewlineReplacement);

            var stem = Path.GetFileNameWithoutExtension(inputPath);
            var result = new ConversionResult { InputPath = inputPath };

            foreach (var sheet in reader.ReadSheets())
            {
                var safeName = SanitizeSheetName(sheet.Name);
                var outputPath = Path.Combine(outputDirectory, $"{stem}_{safeName}.dat");
                _log?.Invoke($"  → sheet \"{sheet.Name}\" → {Path.GetFileName(outputPath)}");

                using var writer = new ConcordanceDatWriter(outputPath, formatter, _options.OutputEncoding);
                var rows = _options.SkipHeader ? sheet.Rows.Skip(1) : sheet.Rows;
                writer.WriteRecords(rows);

                result.Sheets.Add(new SheetResult
                {
                    SheetName = sheet.Name,
                    OutputPath = outputPath,
                    RecordsWritten = writer.RecordsWritten,
                });
                _log?.Invoke($"    wrote {writer.RecordsWritten:N0} record(s)");
            }

            return result;
        }

        private ISpreadsheetReader OpenReader(string inputPath, SpreadsheetFormat format)
        {
            switch (format)
            {
                case SpreadsheetFormat.Csv:
                    return new CsvReader(inputPath, _options.CsvDelimiter);

                case SpreadsheetFormat.OpenXml:
                    return new OpenXmlReader(inputPath);

                case SpreadsheetFormat.LegacyXls:
                case SpreadsheetFormat.OpenDocument:
                case SpreadsheetFormat.AppleNumbers:
                    return OpenViaLibreOffice(inputPath, format);

                default:
                    throw new NotSupportedException(
                        $"Unrecognized spreadsheet format for '{inputPath}'. Supported: .csv, .tsv, .txt, .xlsx, .xlsm, .xls, .ods, .numbers.");
            }
        }

        private LibreOfficeReader OpenViaLibreOffice(string inputPath, SpreadsheetFormat format)
        {
            var soffice = LibreOfficeLocator.Locate(_options.SofficePath);
            if (soffice == null)
            {
                throw new InvalidOperationException(
                    $"{format} input requires LibreOffice. Install LibreOffice and ensure 'soffice' is on PATH, " +
                    "set the SOFFICE_PATH environment variable, or pass --soffice-path.");
            }
            _log?.Invoke($"  using LibreOffice at: {soffice}");
            return new LibreOfficeReader(inputPath, soffice);
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
            var clean = new string(chars).Trim();
            return string.IsNullOrEmpty(clean) ? "Sheet" : clean;
        }
    }
}
