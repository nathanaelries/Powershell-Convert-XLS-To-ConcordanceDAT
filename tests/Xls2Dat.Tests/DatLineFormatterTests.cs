using FluentAssertions;
using Xls2Dat.Core.Formatting;
using Xunit;

namespace Xls2Dat.Tests
{
    public class DatLineFormatterTests
    {
        private const char Pipe = (char)0x14;
        private const char Thorn = (char)0xFE;
        private const char Reg = (char)0xAE;

        [Fact]
        public void Three_simple_fields_are_wrapped_and_joined()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { "a", "b", "c" });
            result.Should().Be($"{Thorn}a{Thorn}{Pipe}{Thorn}b{Thorn}{Pipe}{Thorn}c{Thorn}");
        }

        [Fact]
        public void Empty_fields_round_trip_as_empty_qualified_pairs()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new string?[] { "a", "", null, "d" });
            result.Should().Be($"{Thorn}a{Thorn}{Pipe}{Thorn}{Thorn}{Pipe}{Thorn}{Thorn}{Pipe}{Thorn}d{Thorn}");
        }

        [Fact]
        public void Embedded_newline_becomes_registered_sign()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { "line1\nline2" });
            result.Should().Be($"{Thorn}line1{Reg}line2{Thorn}");
        }

        [Fact]
        public void Crlf_collapses_to_single_replacement()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { "line1\r\nline2" });
            result.Should().Be($"{Thorn}line1{Reg}line2{Thorn}");
        }

        [Fact]
        public void Embedded_qualifier_is_stripped()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { $"foo{Thorn}bar" });
            result.Should().Be($"{Thorn}foobar{Thorn}");
        }

        [Fact]
        public void Embedded_delimiter_is_preserved_verbatim()
        {
            // Delimiter inside a field is fine — the qualifier wrappers disambiguate.
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { $"foo{Pipe}bar" });
            result.Should().Be($"{Thorn}foo{Pipe}bar{Thorn}");
        }

        [Fact]
        public void Unicode_is_passed_through()
        {
            var formatter = new DatLineFormatter();
            var result = formatter.FormatRecord(new[] { "café", "naïve", "日本語" });
            result.Should().Be($"{Thorn}café{Thorn}{Pipe}{Thorn}naïve{Thorn}{Pipe}{Thorn}日本語{Thorn}");
        }

        [Fact]
        public void Single_field_emits_qualified_value()
        {
            var formatter = new DatLineFormatter();
            formatter.FormatRecord(new[] { "x" }).Should().Be($"{Thorn}x{Thorn}");
        }

        [Fact]
        public void Empty_record_produces_empty_string()
        {
            var formatter = new DatLineFormatter();
            formatter.FormatRecord(System.Array.Empty<string?>()).Should().Be(string.Empty);
        }

        [Fact]
        public void Constructor_rejects_matching_delimiter_and_qualifier()
        {
            var act = () => new DatLineFormatter(fieldDelimiter: 'X', textQualifier: 'X');
            act.Should().Throw<System.ArgumentException>();
        }

        [Fact]
        public void Custom_delimiters_are_honored()
        {
            var formatter = new DatLineFormatter('|', '"', '~');
            formatter.FormatRecord(new[] { "a", "b\nc" }).Should().Be("\"a\"|\"b~c\"");
        }
    }
}
