using System.IO;
using System.Linq;
using FluentAssertions;
using Xls2Dat.Core.Readers;
using Xunit;

namespace Xls2Dat.Tests
{
    public class CsvReaderTests
    {
        [Fact]
        public void Reads_csv_with_inferred_delimiter()
        {
            var path = TempFile("a,b,c\n1,2,3\n4,5,6", ".csv");
            try
            {
                using var reader = new CsvReader(path);
                var sheets = reader.ReadSheets().ToList();
                sheets.Should().HaveCount(1);
                var rows = sheets[0].Rows.ToList();
                rows.Should().HaveCount(3);
                rows[0].Should().Equal("a", "b", "c");
                rows[2].Should().Equal("4", "5", "6");
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Reads_tsv_with_tab_delimiter()
        {
            var path = TempFile("a\tb\tc\n1\t2\t3", ".tsv");
            try
            {
                using var reader = new CsvReader(path);
                var rows = reader.ReadSheets().Single().Rows.ToList();
                rows[0].Should().Equal("a", "b", "c");
                rows[1].Should().Equal("1", "2", "3");
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Handles_quoted_field_with_embedded_comma_and_newline()
        {
            var path = TempFile("a,\"b,1\nb2\",c\n1,2,3", ".csv");
            try
            {
                using var reader = new CsvReader(path);
                var rows = reader.ReadSheets().Single().Rows.ToList();
                rows.Should().HaveCount(2);
                rows[0].Should().Equal("a", "b,1\nb2", "c");
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Sheet_name_is_filename_stem()
        {
            var path = TempFile("a,b", ".csv");
            try
            {
                using var reader = new CsvReader(path);
                reader.ReadSheets().Single().Name
                    .Should().Be(Path.GetFileNameWithoutExtension(path));
            }
            finally { File.Delete(path); }
        }

        private static string TempFile(string content, string ext)
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}{ext}");
            File.WriteAllText(path, content);
            return path;
        }
    }
}
