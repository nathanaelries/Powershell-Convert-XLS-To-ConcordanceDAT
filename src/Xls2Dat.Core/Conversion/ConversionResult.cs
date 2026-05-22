using System.Collections.Generic;

namespace Xls2Dat.Core.Conversion
{
    public sealed class ConversionResult
    {
        public string InputPath { get; init; } = "";
        public List<SheetResult> Sheets { get; } = new();
    }

    public sealed class SheetResult
    {
        public string SheetName { get; init; } = "";
        public string OutputPath { get; init; } = "";
        public long RecordsWritten { get; init; }
    }
}
