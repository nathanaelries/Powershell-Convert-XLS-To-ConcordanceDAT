using System.Text;
using Xls2Dat.Core.Formatting;

namespace Xls2Dat.Core.Conversion
{
    public sealed class ConversionOptions
    {
        public char FieldDelimiter { get; set; } = DatLineFormatter.DefaultFieldDelimiter;
        public char TextQualifier { get; set; } = DatLineFormatter.DefaultTextQualifier;
        public char NewlineReplacement { get; set; } = DatLineFormatter.DefaultNewlineReplacement;
        public Encoding? OutputEncoding { get; set; } // null => UTF-16 LE w/ BOM
        public bool SkipHeader { get; set; }
        public string? SofficePath { get; set; }
        public string? CsvDelimiter { get; set; } // null => infer
    }
}
