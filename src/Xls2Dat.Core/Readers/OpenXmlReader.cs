using System.Collections.Generic;
using ClosedXML.Excel;

namespace Xls2Dat.Core.Readers
{
    public sealed class OpenXmlReader : ISpreadsheetReader
    {
        private readonly XLWorkbook _workbook;

        public OpenXmlReader(string path)
        {
            _workbook = new XLWorkbook(path);
        }

        public IEnumerable<Sheet> ReadSheets()
        {
            foreach (var sheet in _workbook.Worksheets)
            {
                yield return new Sheet(sheet.Name, ReadRows(sheet));
            }
        }

        private static IEnumerable<string?[]> ReadRows(IXLWorksheet sheet)
        {
            var range = sheet.RangeUsed();
            if (range == null) yield break;

            int lastColumn = range.LastColumn().ColumnNumber();
            foreach (var row in range.RowsUsed())
            {
                var fields = new string?[lastColumn];
                for (int c = 1; c <= lastColumn; c++)
                {
                    var cell = row.Cell(c);
                    fields[c - 1] = cell.IsEmpty() ? null : cell.GetFormattedString();
                }
                yield return fields;
            }
        }

        public void Dispose() => _workbook.Dispose();
    }
}
