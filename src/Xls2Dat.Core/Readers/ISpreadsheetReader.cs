using System;
using System.Collections.Generic;

namespace Xls2Dat.Core.Readers
{
    public interface ISpreadsheetReader : IDisposable
    {
        IEnumerable<Sheet> ReadSheets();
    }

    public sealed class Sheet
    {
        public string Name { get; }
        public IEnumerable<string?[]> Rows { get; }

        public Sheet(string name, IEnumerable<string?[]> rows)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
        }
    }
}
