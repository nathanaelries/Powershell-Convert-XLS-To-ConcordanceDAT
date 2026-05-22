using System.IO;
using FluentAssertions;
using Xls2Dat.Core.Detection;
using Xunit;

namespace Xls2Dat.Tests
{
    public class FormatDetectorTests
    {
        [Theory]
        [InlineData("a.csv", SpreadsheetFormat.Csv)]
        [InlineData("a.tsv", SpreadsheetFormat.Csv)]
        [InlineData("a.txt", SpreadsheetFormat.Csv)]
        [InlineData("a.xlsx", SpreadsheetFormat.OpenXml)]
        [InlineData("a.xlsm", SpreadsheetFormat.OpenXml)]
        [InlineData("a.xls", SpreadsheetFormat.LegacyXls)]
        [InlineData("a.ods", SpreadsheetFormat.OpenDocument)]
        [InlineData("a.numbers", SpreadsheetFormat.AppleNumbers)]
        public void Detects_by_extension_when_magic_bytes_are_absent(string filename, SpreadsheetFormat expected)
        {
            var path = WriteTempFile(filename, "a,b,c\n1,2,3");
            try
            {
                // For non-CSV extensions with text content, the magic-byte check returns Unknown
                // and the detector should fall back to the extension.
                FormatDetector.Detect(path).Should().Be(expected);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Detects_zip_container_as_openxml_when_extension_is_xlsx()
        {
            var path = WriteTempBytes("a.xlsx", new byte[] { 0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0 });
            try { FormatDetector.Detect(path).Should().Be(SpreadsheetFormat.OpenXml); }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Detects_ole_compound_as_legacy_xls()
        {
            var path = WriteTempBytes("a.xls", new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
            try { FormatDetector.Detect(path).Should().Be(SpreadsheetFormat.LegacyXls); }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Throws_when_file_missing()
        {
            var act = () => FormatDetector.Detect(Path.Combine(Path.GetTempPath(), "does-not-exist.xlsx"));
            act.Should().Throw<FileNotFoundException>();
        }

        private static string WriteTempFile(string name, string content)
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}-{name}");
            File.WriteAllText(path, content);
            return path;
        }

        private static string WriteTempBytes(string name, byte[] bytes)
        {
            var path = Path.Combine(Path.GetTempPath(), $"xls2dat-test-{System.Guid.NewGuid():N}-{name}");
            File.WriteAllBytes(path, bytes);
            return path;
        }
    }
}
