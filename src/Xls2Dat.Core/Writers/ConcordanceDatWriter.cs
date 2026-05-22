using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xls2Dat.Core.Formatting;

namespace Xls2Dat.Core.Writers
{
    public sealed class ConcordanceDatWriter : IDisposable
    {
        private readonly StreamWriter _writer;
        private readonly DatLineFormatter _formatter;

        public long RecordsWritten { get; private set; }

        public ConcordanceDatWriter(string outputPath, DatLineFormatter formatter, Encoding? encoding = null)
        {
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir!);

            encoding ??= new UnicodeEncoding(bigEndian: false, byteOrderMark: true);
            _writer = new StreamWriter(outputPath, append: false, encoding, bufferSize: 65536)
            {
                NewLine = "\r\n",
                AutoFlush = false,
            };
            _formatter = formatter;
        }

        public void WriteRecords(IEnumerable<string?[]> records)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            foreach (var record in records)
            {
                if (record == null || record.Length == 0) continue;
                _writer.WriteLine(_formatter.FormatRecord(record));
                RecordsWritten++;
            }
        }

        public void Dispose()
        {
            _writer.Flush();
            _writer.Dispose();
        }
    }
}
