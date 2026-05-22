using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;

namespace Xls2Dat.Core.Readers
{
    public sealed class CsvReader : ISpreadsheetReader
    {
        private readonly string _path;
        private readonly string? _delimiter;

        public CsvReader(string path, string? delimiter = null)
        {
            _path = path;
            _delimiter = delimiter;
        }

        public IEnumerable<Sheet> ReadSheets()
        {
            yield return new Sheet(Path.GetFileNameWithoutExtension(_path), ReadRows());
        }

        private IEnumerable<string?[]> ReadRows()
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = false,
                Delimiter = _delimiter ?? InferDelimiter(_path),
                BadDataFound = null,
                MissingFieldFound = null,
                IgnoreBlankLines = true,
            };

            using var reader = new StreamReader(_path, detectEncodingFromByteOrderMarks: true);
            using var csv = new CsvParser(reader, config);

            while (csv.Read())
            {
                var record = csv.Record;
                if (record != null) yield return record;
            }
        }

        public void Dispose() { }

        private static string InferDelimiter(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".tsv" => "\t",
                ".txt" => "\t",
                _ => ","
            };
        }
    }
}
