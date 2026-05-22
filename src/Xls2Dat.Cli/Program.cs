using System;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Globalization;
using System.IO;
using System.Text;
using Xls2Dat.Core.Conversion;

namespace Xls2Dat.Cli
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            var inputArg = new Argument<FileInfo>("input", "Spreadsheet to convert (.csv, .tsv, .txt, .xlsx, .xlsm, .xls, .ods, .numbers).");
            var outputOpt = new Option<DirectoryInfo>(new[] { "--output", "-o" }, "Directory where .dat files will be written.") { IsRequired = true };
            var fieldDelimOpt = new Option<string>("--field-delimiter", () => "0x14", "Field delimiter (char or 0xNN).");
            var textQualOpt = new Option<string>("--text-qualifier", () => "0xFE", "Text qualifier (char or 0xNN).");
            var newlineReplOpt = new Option<string>("--newline-replacement", () => "0xAE", "Replacement for embedded newlines (char or 0xNN).");
            var encodingOpt = new Option<string>("--encoding", () => "utf-16le", "Output encoding: utf-16le, utf-16be, utf-8, utf-8-bom.");
            var skipHeaderOpt = new Option<bool>("--skip-header", "Drop the first row of each sheet.");
            var sofficeOpt = new Option<string?>("--soffice-path", "Explicit path to LibreOffice soffice executable.");
            var csvDelimOpt = new Option<string?>("--csv-delimiter", "Override CSV/TSV delimiter (e.g. ',', ';', '\\t').");
            var verboseOpt = new Option<bool>(new[] { "--verbose", "-v" }, "Verbose logging.");

            var root = new RootCommand("Convert spreadsheets (.csv, .xlsx, .xls, .ods, .numbers, ...) to Concordance .dat files.")
            {
                inputArg, outputOpt, fieldDelimOpt, textQualOpt, newlineReplOpt,
                encodingOpt, skipHeaderOpt, sofficeOpt, csvDelimOpt, verboseOpt,
            };

            root.SetHandler((InvocationContext ctx) =>
            {
                var p = ctx.ParseResult;
                var input = p.GetValueForArgument(inputArg);
                var output = p.GetValueForOption(outputOpt)!;
                var verbose = p.GetValueForOption(verboseOpt);

                try
                {
                    var options = new ConversionOptions
                    {
                        FieldDelimiter = ParseChar(p.GetValueForOption(fieldDelimOpt)!, nameof(fieldDelimOpt)),
                        TextQualifier = ParseChar(p.GetValueForOption(textQualOpt)!, nameof(textQualOpt)),
                        NewlineReplacement = ParseChar(p.GetValueForOption(newlineReplOpt)!, nameof(newlineReplOpt)),
                        OutputEncoding = ParseEncoding(p.GetValueForOption(encodingOpt)!),
                        SkipHeader = p.GetValueForOption(skipHeaderOpt),
                        SofficePath = p.GetValueForOption(sofficeOpt),
                        CsvDelimiter = p.GetValueForOption(csvDelimOpt),
                    };

                    Action<string>? log = verbose ? Console.Error.WriteLine : null;
                    var converter = new SpreadsheetConverter(options, log);
                    var result = converter.Convert(input.FullName, output.FullName);

                    long total = 0;
                    foreach (var sheet in result.Sheets) total += sheet.RecordsWritten;
                    Console.Out.WriteLine($"Wrote {result.Sheets.Count} sheet(s), {total:N0} record(s) to {output.FullName}");
                    ctx.ExitCode = 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"error: {ex.Message}");
                    if (verbose) Console.Error.WriteLine(ex);
                    ctx.ExitCode = 1;
                }
            });

            return root.Invoke(args);
        }

        private static char ParseChar(string s, string optName)
        {
            if (string.IsNullOrEmpty(s))
                throw new ArgumentException($"{optName} cannot be empty.");

            if (s.StartsWith("0x", StringComparison.OrdinalIgnoreCase) || s.StartsWith("\\x", StringComparison.OrdinalIgnoreCase))
            {
                var hex = s.Substring(2);
                if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code))
                    return (char)code;
                throw new ArgumentException($"{optName}: invalid hex value '{s}'.");
            }

            if (s.Length == 1) return s[0];
            throw new ArgumentException($"{optName}: expected a single character or 0xNN, got '{s}'.");
        }

        private static Encoding ParseEncoding(string s) => s.ToLowerInvariant() switch
        {
            "utf-16le" or "utf16le" or "unicode" => new UnicodeEncoding(bigEndian: false, byteOrderMark: true),
            "utf-16be" or "utf16be" => new UnicodeEncoding(bigEndian: true, byteOrderMark: true),
            "utf-8" or "utf8" => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            "utf-8-bom" or "utf8-bom" => new UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
            _ => throw new ArgumentException($"Unsupported encoding '{s}'. Use utf-16le, utf-16be, utf-8, or utf-8-bom."),
        };
    }
}
