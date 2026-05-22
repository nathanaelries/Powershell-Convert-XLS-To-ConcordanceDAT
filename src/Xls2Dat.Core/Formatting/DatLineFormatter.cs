using System;
using System.Text;

namespace Xls2Dat.Core.Formatting
{
    public sealed class DatLineFormatter
    {
        public const char DefaultFieldDelimiter = (char)0x14;     // DC4
        public const char DefaultTextQualifier = (char)0xFE;      // thorn
        public const char DefaultNewlineReplacement = (char)0xAE; // registered

        private readonly char _delimiter;
        private readonly char _qualifier;
        private readonly char _newlineReplacement;
        private readonly string _between;

        public DatLineFormatter(
            char fieldDelimiter = DefaultFieldDelimiter,
            char textQualifier = DefaultTextQualifier,
            char newlineReplacement = DefaultNewlineReplacement)
        {
            if (fieldDelimiter == textQualifier)
                throw new ArgumentException("Field delimiter and text qualifier must differ.");

            _delimiter = fieldDelimiter;
            _qualifier = textQualifier;
            _newlineReplacement = newlineReplacement;
            _between = new string(new[] { _qualifier, _delimiter, _qualifier });
        }

        public string FormatRecord(string?[] fields)
        {
            if (fields == null) throw new ArgumentNullException(nameof(fields));
            if (fields.Length == 0) return string.Empty;

            var sb = new StringBuilder(64 * fields.Length);
            sb.Append(_qualifier);
            AppendField(sb, fields[0]);
            for (int i = 1; i < fields.Length; i++)
            {
                sb.Append(_between);
                AppendField(sb, fields[i]);
            }
            sb.Append(_qualifier);
            return sb.ToString();
        }

        private void AppendField(StringBuilder sb, string? value)
        {
            if (string.IsNullOrEmpty(value)) return;

            for (int i = 0; i < value!.Length; i++)
            {
                char c = value[i];
                if (c == _qualifier) continue;
                if (c == '\r') continue;
                if (c == '\n') { sb.Append(_newlineReplacement); continue; }
                sb.Append(c);
            }
        }
    }
}
