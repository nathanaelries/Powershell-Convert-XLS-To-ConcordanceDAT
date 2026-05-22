using System.IO;
using System.Text;
using FluentAssertions;
using Xls2Dat.Core.Formatting;
using Xls2Dat.Core.Writers;
using Xunit;

namespace Xls2Dat.Tests
{
    public class ConcordanceDatWriterTests
    {
        [Fact]
        public void Writes_utf16le_with_bom_by_default()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}.dat");
            try
            {
                using (var writer = new ConcordanceDatWriter(path, new DatLineFormatter()))
                {
                    writer.WriteRecords(new[] { new string?[] { "a", "b" } });
                }

                var bytes = File.ReadAllBytes(path);
                bytes[0].Should().Be(0xFF);
                bytes[1].Should().Be(0xFE);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Writes_crlf_line_terminators()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}.dat");
            try
            {
                using (var writer = new ConcordanceDatWriter(path, new DatLineFormatter(), new UTF8Encoding(false)))
                {
                    writer.WriteRecords(new[]
                    {
                        new string?[] { "a" },
                        new string?[] { "b" },
                    });
                }

                var text = File.ReadAllText(path, Encoding.UTF8);
                text.Should().Contain("\r\n");
                text.Should().NotContain("\n\n");
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Skips_null_and_empty_records()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}.dat");
            try
            {
                using var writer = new ConcordanceDatWriter(path, new DatLineFormatter());
                writer.WriteRecords(new[]
                {
                    new string?[] { "a" },
                    System.Array.Empty<string?>(),
                    new string?[] { "b" },
                });
                writer.RecordsWritten.Should().Be(2);
            }
            finally { File.Delete(path); }
        }
    }
}
