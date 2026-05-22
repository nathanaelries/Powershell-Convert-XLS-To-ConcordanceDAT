using System;
using System.IO;

namespace Xls2Dat.Core.Detection
{
    public static class FormatDetector
    {
        public static SpreadsheetFormat Detect(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path required.", nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException("Input file not found.", path);

            var byExt = ByExtension(path);
            var byMagic = ByMagicBytes(path);

            // Magic-byte signals win for container-based formats (xlsx/ods/numbers are all zip archives,
            // .xls is OLE Compound Document). Extension wins for plain-text CSV variants.
            if (byMagic != SpreadsheetFormat.Unknown && byExt != SpreadsheetFormat.Csv)
                return byMagic;

            return byExt != SpreadsheetFormat.Unknown ? byExt : byMagic;
        }

        private static SpreadsheetFormat ByExtension(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".csv" or ".tsv" or ".txt" => SpreadsheetFormat.Csv,
                ".xlsx" or ".xlsm" => SpreadsheetFormat.OpenXml,
                ".xls" => SpreadsheetFormat.LegacyXls,
                ".ods" => SpreadsheetFormat.OpenDocument,
                ".numbers" => SpreadsheetFormat.AppleNumbers,
                _ => SpreadsheetFormat.Unknown
            };
        }

        private static SpreadsheetFormat ByMagicBytes(string path)
        {
            var head = new byte[8];
            int read;
            using (var fs = File.OpenRead(path))
                read = fs.Read(head, 0, head.Length);

            if (read < 4) return SpreadsheetFormat.Unknown;

            // OLE Compound Document (legacy .xls): D0 CF 11 E0 A1 B1 1A E1
            if (read >= 8 &&
                head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0 &&
                head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1)
                return SpreadsheetFormat.LegacyXls;

            // Zip-based container (xlsx / ods / numbers): "PK\x03\x04"
            if (head[0] == 0x50 && head[1] == 0x4B && head[2] == 0x03 && head[3] == 0x04)
            {
                // Disambiguate by inspecting the zip's mimetype/content. Use extension as the
                // heuristic — if the user lied about the extension we'll still try OpenXml first
                // and fall back to LibreOffice.
                var ext = Path.GetExtension(path).ToLowerInvariant();
                return ext switch
                {
                    ".ods" => SpreadsheetFormat.OpenDocument,
                    ".numbers" => SpreadsheetFormat.AppleNumbers,
                    _ => SpreadsheetFormat.OpenXml
                };
            }

            return SpreadsheetFormat.Unknown;
        }
    }
}
